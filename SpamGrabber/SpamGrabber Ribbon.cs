using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;
using SpamGrabberControl;

namespace SpamGrabber
{
    public partial class SpamGrabber_Ribbon
    {
        #region Button Event Handlers

        private void btnReportDefaultSpam_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(SpamGrabberCommon.GlobalSettings.DefaultSpamProfileId))
            {
                MessageBox.Show("You have not yet set a default spam profile. Please open the SpamGrabber settings dialog and set a default spam profile");
                return;
            }
            SendReports(SpamGrabberCommon.GlobalSettings.DefaultSpamProfileId);
        }

        private void btnCopyToClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer exp = Globals.ThisAddIn.Application.ActiveExplorer();
            if (exp.Selection.Count > 0)
            {
                Clipboard.SetText(GetMessageSource((MailItem)exp.Selection[1], false));
            }
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                frmOptions myOptions = new frmOptions();
                myOptions.ShowDialog();
                if (myOptions.DialogResult == DialogResult.OK)
                {
                    // Refresh the command bar
                    this.ShowHideButtons();
                }
            }
            catch (System.Exception ex) // TODO we should not catch all exceptions
            {
                MessageBox.Show("caught: \r\n" + ex.ToString());
            }
        }

        private void btnReportCustom_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.ddlReportTo.SelectedItem != null)
            {
                this.SendReports(this.ddlReportTo.SelectedItem.Tag.ToString());
            }
        }

        private void btnSafeView_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer exp = Globals.ThisAddIn.Application.ActiveExplorer();
            if (exp.Selection.Count > 0)
            {
                frmPreview objPreview = new frmPreview();
                objPreview.ClearItems();
                foreach (object objItem in exp.Selection)
                {
                    if (objItem is MailItem || objItem is PostItem)
                        objPreview.Items.Add(objItem);
                }
                objPreview.ShowDialog();
            }
        }

        #endregion

        #region Common Functions

        private void SendReports(string profileID)
        {
            SpamGrabberCommon.Profile profile = new SpamGrabberCommon.Profile(profileID);

            if (profile.AskVerify)
            {
                if (MessageBox.Show("Are you sure you want to report the selected item(s)?", "Report messages", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
            }

            Explorer exp = Globals.ThisAddIn.Application.ActiveExplorer();

            // Create the report email
            MailItem reportEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
            reportEmail.Subject = profile.ReportSubject;
            foreach (string toAddress in profile.ToAddresses)
            {
                reportEmail.To += toAddress + ";";
            }
            foreach (string bccAddress in profile.BccAddresses)
            {
                reportEmail.BCC += bccAddress + ";";
            }
            reportEmail.BodyFormat = OlBodyFormat.olFormatPlain;
            reportEmail.Body = profile.MessageBody;

            // Make sure at least one item is sent
            bool bItemsSelected = false;

            foreach (Object item in exp.Selection)
            {
                if (item is MailItem){
                    bItemsSelected = true;
                    MailItem mail = (MailItem)item;
                    if (profile.UseRFC)
                    {
                        reportEmail.Attachments.Add(mail, OlAttachmentType.olEmbeddeditem);
                    }
                    else
                    {
                        // Create temp file
                        string fileName = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".txt");
                        TextWriter tw = new StreamWriter(fileName);
                        tw.Write(GetMessageSource(mail, profile.CleanHeaders));
                        tw.Close();
                        reportEmail.Attachments.Add(fileName);
                    }
                }
            }
            
            if (bItemsSelected)
            {
                // Do we need to keep a copy
                if (!profile.KeepCopy)
                {
                    reportEmail.DeleteAfterSubmit = true;
                }

                // Send the report
                reportEmail.Send();

                // Sort out actions on the source emails
                foreach (Object item in exp.Selection)
                {
                    if (item is MailItem)
                    {
                        MailItem mail = (MailItem)item;
                        if (profile.MarkAsReadAfterReport)
                        {
                            mail.UnRead = false;
                        }
                        if (profile.DeleteAfterReport)
                        {
                            mail.Delete();
                        }
                        else if (profile.MoveToFolderAfterReport)
                        {
                            mail.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetFolderFromID(
                                profile.MoveFolderName, profile.MoveFolderStoreId));
                        }
                    }
                }

                // Send / receive

            }
            else
            {
                reportEmail.Delete();
            }
        }

        private void SpamGrabber_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            foreach (SpamGrabberCommon.Profile profile in SpamGrabberCommon.UserProfiles.ProfileList)
            {
                //SpamGrabberCommon.Profile profile = new SpamGrabberCommon.Profile(profileID);
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = profile.Name;
                item.Tag = profile.Id;
                this.ddlReportTo.Items.Add(item);
            }
            // Show / hide buttons based on settings
            this.ShowHideButtons();
        }

        private string GetMessageSource(MailItem message, bool cleanHeaders)
        {
            string headers = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");
            return string.Format("{1}{2}", Environment.NewLine,
                cleanHeaders ? SpamGrabberCommon.SGGlobals.RepairHeaders(headers, message.BodyFormat.Equals(OlBodyFormat.olFormatHTML)) : headers,
                message.BodyFormat == OlBodyFormat.olFormatHTML ? message.HTMLBody : message.Body);
        }

        private void ShowHideButtons()
        {
            this.btnReportDefaultHam.Visible = SpamGrabberCommon.GlobalSettings.ShowHamButton;
            this.btnCopyToClipboard.Visible = SpamGrabberCommon.GlobalSettings.ShowCopyButton;
            this.btnSafeView.Visible = SpamGrabberCommon.GlobalSettings.ShowPreviewButton;
            this.gpSettings.Visible = SpamGrabberCommon.GlobalSettings.ShowSettingsButton;
            this.boxReportTo.Visible = SpamGrabberCommon.GlobalSettings.ShowSelectButton;
        }

        #endregion
    }
}
