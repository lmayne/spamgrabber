using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;

namespace SpamGrabberCommon
{
    public class Reporting
    {
        private static Microsoft.Office.Interop.Outlook.Application _app;
        public static Microsoft.Office.Interop.Outlook.Application Application
        {
            get
            {
                return _app;
            }
            set
            {
                _app = value;
            }
        }

        public static void SendReports(string profileID)
        {
            SpamGrabberCommon.Profile profile = new SpamGrabberCommon.Profile(profileID);

            if (profile.AskVerify)
            {
                if (MessageBox.Show("Are you sure you want to report the selected item(s)?", "Report messages", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
            }

            Explorer exp = _app.Application.ActiveExplorer();

            // Create a collection to hold references to the attachments
            List<string> attachmentFiles = new List<string>();

            // Make sure at least one item is sent
            bool bItemsSelected = false;

            // First make sure the selected emails have been downloaded
            bool bNeedsSendReceive = false;
            for (int i = 1; i <= exp.Selection.Count; i++)
            {
                if (exp.Selection[i] is MailItem)
                {
                    MailItem mail = (MailItem)exp.Selection[i];
                    bItemsSelected = true;
                    // If the item has not been downloaded, mark for download
                    if (mail.DownloadState == OlDownloadState.olHeaderOnly)
                    {
                        bNeedsSendReceive = true;
                        mail.MarkForDownload = OlRemoteStatus.olMarkedForDownload;
                        mail.Save();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                }
            }
            if (bNeedsSendReceive)
            {
                // Download the marked emails
                // TODO: Trying to carry on at this point returns blank email bodies. Try and find a way of downloading them properly.
                _app.Session.SendAndReceive(false);
                MessageBox.Show("One of more emails were not downloaded from the server. Please ensure they are now downloaded and click report again",
                    "SpamGrabber", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (bItemsSelected)
            {
                // Now get references to all the items
                for (int i = 1; i <= exp.Selection.Count; i++)
                {
                    if (exp.Selection[i] is MailItem)
                    {
                        MailItem mail = (MailItem)exp.Selection[i];
                        if (profile.UseRFC)
                        {
                            // Direct attaching seems to be buggy. Save the mailitem first
                            string fileName = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".msg");
                            mail.SaveAs(fileName);
                            attachmentFiles.Add(fileName);
                        }
                        else
                        {
                            // Create temp text file
                            string fileName = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".txt");
                            TextWriter tw = new StreamWriter(fileName);
                            tw.Write(GetMessageSource(mail, profile.CleanHeaders));
                            tw.Close();
                            attachmentFiles.Add(fileName);
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                    }
                }

                // Are we using a single email or one per report?
                if (profile.SendMultiple)
                {
                    // Create the report email
                    MailItem reportEmail = CreateReportEmail(profile);

                    // Attach the files
                    foreach (string attachment in attachmentFiles)
                    {
                        reportEmail.Attachments.Add(attachment);
                    }

                    // Send the report
                    reportEmail.Send();

                    // Do we need to keep a copy?
                    if (!profile.KeepCopy)
                    {
                        reportEmail.Delete();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(reportEmail);
                }
                else
                {
                    // Send one email per report
                    foreach (string attachment in attachmentFiles)
                    {
                        MailItem reportEmail = CreateReportEmail(profile);
                        reportEmail.Attachments.Add(attachment);
                        reportEmail.Send();
                        // Do we need to keep a copy?
                        if (!profile.KeepCopy)
                        {
                            reportEmail.Delete();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(reportEmail);
                    }
                }

                // Sort out actions on the source emails
                for (int i = 1; i <= exp.Selection.Count; i++)
                {
                    if (exp.Selection[i] is MailItem)
                    {
                        MailItem mail = (MailItem)exp.Selection[i];
                        if (profile.MarkAsReadAfterReport)
                        {
                            mail.UnRead = false;
                        }
                        if (profile.DeleteAfterReport)
                        {
                            mail.UnRead = false;
                            mail.Delete();
                        }
                        else if (profile.MoveToFolderAfterReport)
                        {
                            mail.Move(_app.GetNamespace("MAPI").GetFolderFromID(
                                profile.MoveFolderName, profile.MoveFolderStoreId));
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                    }
                }
            }
        }

        public static string GetMessageSource(MailItem message, bool cleanHeaders)
        {
            string headers = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");
            return string.Format("{1}{0}{2}", Environment.NewLine,
                cleanHeaders ? SpamGrabberCommon.SGGlobals.RepairHeaders(headers, message.BodyFormat.Equals(OlBodyFormat.olFormatHTML)) : headers,
                message.BodyFormat == OlBodyFormat.olFormatHTML ? message.HTMLBody : message.Body);
        }

        private static MailItem CreateReportEmail(SpamGrabberCommon.Profile profile)
        {
            // Create the report email
            MailItem reportEmail = (MailItem)_app.CreateItem(OlItemType.olMailItem);
            reportEmail.Subject = profile.ReportSubject;
            string strTo = "";
            foreach (string toAddress in profile.ToAddresses)
            {
                strTo += toAddress + ";";
            }
            reportEmail.To = strTo;

            string strBcc = "";
            foreach (string bccAddress in profile.BccAddresses)
            {
                strBcc += bccAddress + ";";
            }
            reportEmail.BCC = strBcc;

            reportEmail.BodyFormat = OlBodyFormat.olFormatPlain;
            reportEmail.Body = profile.MessageBody;
            return reportEmail;
        }
    }
}
