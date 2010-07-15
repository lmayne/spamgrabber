using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using X4U.Outlook;
using System.IO;
using Redemption;


/*
multipart-body := preamble 1*encapsulation 
               close-delimiter epilogue 

encapsulation := delimiter CRLF body-part 

delimiter := CRLF "--" boundary   ; taken from  Content-Type 
field. 
                               ;   when   content-type    is 
multipart 
                             ; There must be no space 
                             ; between "--" and boundary. 

close-delimiter := delimiter "--" ; Again, no  space  before 
"--" 

preamble :=  *text                  ;  to  be  ignored  upon 
receipt. 

epilogue :=  *text                  ;  to  be  ignored  upon 
receipt. 

body-part = <"message" as defined in RFC 822, 
         with all header fields optional, and with the 
         specified delimiter not occurring anywhere in 
         the message body, either on a line by itself 
         or as a substring anywhere.  Note that the 
         semantics of a part differ from the semantics 
         of a message, as described in the text.> 
*/
namespace SpamGrabberCommon
{
    public class MimeConversions
    {
        public MimeConversions()
        {
            // TODO: add constructor
        }

        /// <summary>
        /// Creates the parsed mime message
        /// </summary>
        /// <param name="header"></param>
        /// <param name="body"></param>
        /// <param name="htmlbody"></param>
        /// <returns></returns>
        private static string ConvertMime(string pstrHeader, string pstrBody, string pstrHtmlbody, string[] pstaAttachments)
        {
            List<string> lstHeader = new List<string>();
            List<string> lstCombined = new List<string>();
            int count = 0;
            //bool isMultiPart = false;
            int attCount = 0;
            bool isAttached = false;

            string boundaryIdentifier = string.Empty;

            string[] del = new string[] { "\r\n" };
            lstHeader.AddRange(pstrHeader.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None));
            if (lstHeader.Count > 0)
                lstHeader[0] = lstHeader[0].Replace("Microsoft Mail Internet Headers Version 2.0\r\n", "");

            Dictionary<string, string> lstBoundaries = GetBoundaryIdTree(lstHeader);

            // remove boundaryparts from attached emails.
            //int nr = 0;
            List<string> orderedHeaders = new List<string>();
            //bool flag = false;
            string[] keys = new string[lstBoundaries.Keys.Count];
            lstBoundaries.Keys.CopyTo(keys, 0);



            foreach (string hs in lstHeader)
            {

                if (count == 0)
                {

                    // process the different parts
                    lstCombined.Add(hs);

                    if (hs.Contains("boundary="))// this is a multipart mime
                    {
                        // capture the boundary tag as we need it for later.
                        int boundaryStart = hs.IndexOf("boundary=") + 9; // length of boundary text
                        int boundaryEnd = hs.IndexOf("\r\n", boundaryStart);
                        boundaryIdentifier = hs.Substring(boundaryStart + 1, boundaryEnd - boundaryStart - 2);

                        //                        isMultiPart = true;
                        lstCombined.AddRange(del);
                        lstCombined.AddRange(del);
                        lstCombined.Add("This is a multi-part message in MIME format.");
                        lstCombined.AddRange(del);
                        lstCombined.AddRange(del);
                    }
                    else// no multipart
                    {
                        if (pstrHtmlbody.Length > 0 && !pstrHtmlbody.Contains("<!-- Converted from text/plain format -->"))
                        {
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                            if (hs.Contains("Content-Transfer-Encoding: quoted-printable"))
                                lstCombined.Add(QuotedPrintable.Encode(pstrHtmlbody));
                            else if (hs.Contains("Content-Transfer-Encoding: base64"))
                                lstCombined.Add(Base64.Encode(pstrHtmlbody));
                            else
                                lstCombined.Add(pstrHtmlbody);


                        }
                        else
                        {
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                            if (hs.Contains("Content-Transfer-Encoding: quoted-printable"))
                                lstCombined.Add(QuotedPrintable.Encode(pstrBody));
                            else if (hs.Contains("Content-Transfer-Encoding: base64"))
                                lstCombined.Add(Base64.Encode(pstrBody));
                            else
                                lstCombined.Add(pstrBody);

                        }

                    }
                }
                else
                {
                    if (hs.Contains("Content-Disposition: attachment;") && hs.Contains(boundaryIdentifier))
                    {

                        lstCombined.Add(hs);
                        isAttached = true;
                        if (pstaAttachments.Length > 0)
                        {
                            // safety check... ensure that we dont get'
                            // index out of bounds error due to faulty headers.
                            if (attCount > pstaAttachments.Length - 1)
                                attCount = pstaAttachments.Length - 1;

                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                            lstCombined.Add(pstaAttachments[attCount]);// add the attachment
                            lstCombined.AddRange(del);
                            attCount++;
                        }
                    }
                    else
                    {
                        lstCombined.Add(hs);

                        if (hs.Contains(boundaryIdentifier) && isAttached)
                            isAttached = false;
                    }

                    //------- - - - - - - - - - - - - - - ------
                    if (!isAttached)
                    {
                        if (hs.Contains("boundary="))
                        {
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                            // add extra line if last char isn't a LF
                            if (!hs.EndsWith("\r\n"))
                            {
                                lstCombined.AddRange(del);
                            }

                        }

                        if (hs.Contains("Content-Type: text/html;"))
                        {
                            if (pstrHtmlbody.Length > 0 && !pstrHtmlbody.Contains("<!-- Converted from text/plain format -->"))
                            {
                                lstCombined.AddRange(del);
                                lstCombined.AddRange(del);

                                if (hs.Contains("Content-Transfer-Encoding: quoted-printable"))
                                    lstCombined.Add(QuotedPrintable.Encode(pstrHtmlbody));
                                else if (hs.Contains("Content-Transfer-Encoding: base64"))
                                    lstCombined.Add(Base64.Encode(pstrHtmlbody));
                                else
                                    lstCombined.Add(pstrHtmlbody);
                            }

                        }
                        else if (hs.Contains("Content-Type: text/plain;"))
                        {
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);// måske ?
                            if (hs.Contains("Content-Transfer-Encoding: quoted-printable"))
                                lstCombined.Add(QuotedPrintable.Encode(pstrBody));
                            else if (hs.Contains("Content-Transfer-Encoding: base64"))
                                lstCombined.Add(Base64.Encode(pstrBody));
                            else
                                lstCombined.Add(pstrBody);

                            lstCombined.AddRange(del);

                        }
                        else if (hs.Contains("Content-Type:"))
                        {
                            //Processed content types:
                            // "image" / "audio" / "video" / "application"
                            if (!hs.Contains("Content-Type: multipart/"))
                            {
                                if (pstaAttachments.Length > 0)
                                {
                                    // safety check... ensure that we dont get'
                                    // index out of bounds error due to faulty headers.
                                    if (attCount > pstaAttachments.Length - 1)
                                        attCount = pstaAttachments.Length - 1;

                                    lstCombined.AddRange(del);
                                    lstCombined.AddRange(del);
                                    lstCombined.Add(pstaAttachments[attCount]);// add the attachment
                                    lstCombined.AddRange(del);
                                    attCount++;
                                }
                            }


                        }
                        else if (hs.EndsWith("--"))
                        {
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                            lstCombined.AddRange(del);
                        }
                        else
                        {
                            if (!hs.Contains("boundary="))
                                lstCombined.AddRange(del);
                        }
                    }

                }

                count++;
            }

            return string.Join("", lstCombined.ToArray());// return something

        }

        /// <summary>
        /// This function builds a parent/child relationship tree from the boundary identifiers
        /// </summary>
        /// <param name="lstHeader">The outlook transport header</param>
        /// <returns>A Dictionary containing the keys as child key, parent value</returns>
        private static Dictionary<string, string> GetBoundaryIdTree(List<string> lstHeader)
        {
            // create boundary tree
            Dictionary<string, string> lstBoundaries = new Dictionary<string, string>();
            string boundaryID = "";
            string parentID = "";
            foreach (string line in lstHeader)
            {
                try
                {
                    //Content-Type: multipart/mixed;
                    //boundary="----=_NextPart_000_000F_01C766E2.4C3E1E90"
                    if (line.Contains("boundary="))
                    {
                        boundaryID = GetBoundaryId(line);
                        if (parentID == "")
                        {
                            lstBoundaries.Add(boundaryID, "");
                            parentID = boundaryID;
                        }
                        //------=_NextPart_000_000F_01C766E2.4C3E1E90
                        //Content-Type: multipart/alternative;
                        //  boundary="----=_NextPart_001_0010_01C766E2.4C3E6CB0"
                        else if (line.Contains(parentID))
                        {
                            if (!lstBoundaries.ContainsKey(boundaryID))
                                lstBoundaries.Add(boundaryID, parentID);
                        }
                        else
                        {
                            parentID = "";
                            lstBoundaries.Add(boundaryID, parentID);
                            parentID = boundaryID;
                        }

                    }

                }
                catch (Exception parseexc)
                {
                    //TODO : Remove for silent falure
                    System.Windows.Forms.MessageBox.Show("Caught: \r\n" + parseexc.ToString());
                }
            }
            return lstBoundaries;
        }

        /// <summary>
        /// Get the boundary id for the supplied line
        /// </summary>
        /// <param name="line">A header line from the transport header.</param>
        /// <returns>The boundary identifier</returns>
        private static string GetBoundaryId(string strLine)
        {
            if (strLine.IndexOf("boundary=") == -1) // no boundary
                return "";

            // capture the boundary tag as we need it for later.
            int bStart = strLine.IndexOf("boundary=") + 9; // length of boundary text
            int bEnd = strLine.IndexOf("\r\n", bStart);
            string boundaryID;
            if (bEnd == -1) // no end delimiter
            {
                if ((strLine.Length - bStart - 2) > 0)
                    boundaryID = strLine.Substring(bStart + 1, strLine.Length - bStart - 2);
                else
                    boundaryID = "";
            }
            else
            {
                if ((bEnd - bStart - 2) > 0)
                boundaryID = strLine.Substring(bStart + 1, bEnd - bStart - 2);
                else
                    boundaryID="";

            }
            return boundaryID;
        }

        /// <summary>
        /// Exports an Outlook MailItem to a file for later attachment.
        /// </summary>
        /// <param name="Message">The message to be exported</param>
        /// <param name="isSpamReport">is it a spamreport</param>
        /// <param name="objDefaultProfile">The reporting profile</param>
        /// <returns>The filename and path as a string for later use.</returns>
        public static string ExportToFile(object pobjMessage, SGGlobals.ReportAction pAction, Profile pobjProfile)
        {
            string filename = string.Empty;
            string path = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache);//the temp cache
            object obj = Activator.CreateInstance(Type.GetTypeFromProgID("safemail.safemailMailItem"));
            Redemption.SafeMailItem safeMail = (Redemption.SafeMailItem)obj;//new SafeMailItem();

            if (pobjMessage is Outlook.MailItem)
                safeMail.Item = pobjMessage;
            else if (pobjMessage is Outlook.PostItem)
                safeMail.Item = pobjMessage;


            //filename = Message.ReceivedTime.ToFileTime() + ".eml";
            filename = SpamGrapperUtility.CreateAttachmentName(pAction, pobjProfile.IncludeIPAddress);
            FileInfo fi = new FileInfo(path + "\\" + filename);

            //Create a file to write to.
            using (StreamWriter sw = fi.CreateText())
            {

                sw.Write(ConvertMail(pobjMessage, pobjProfile));
                sw.Close();

            }
            // returning the filename for caller to use
            // caller must delete the file after use.
            return path + "\\" + filename;
        }

        /// <summary>
        /// Returns an Outlook MailItem as a MAPI string.
        /// </summary>
        /// <param name="Message">The message to be Converted</param>
        /// <returns>The Mail</returns>
        public static string ConvertMail(object pobjMessage, Profile pobjProfile)
        {
            object obj = Activator.CreateInstance(Type.GetTypeFromProgID("safemail.safemailMailItem"));
            Redemption.SafeMailItem safeMail = (Redemption.SafeMailItem)obj;//new SafeMailItem();
            //Redemption.SafeMailItem safeMail = new SafeMailItem();
            if (pobjMessage is Outlook.MailItem)
                safeMail.Item = pobjMessage;
            else if(pobjMessage is Outlook.PostItem)
                safeMail.Item = pobjMessage;

            bool msgIsHTML = false;
            string completeMail = string.Empty;


            if (pobjProfile.FixMIME)
            {
                if (IsValidMIMEHeader(safeMail))
                {
                    completeMail = (GetMimeMail(safeMail));
                }
                else // this is a copy of the part for non FixMIME messages... was thinking of adding the attachments,
                // but I don't really se how :-(
                {
                    string sFilePath = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache) + @"\mailexport.txt";
                    safeMail.SaveAs(sFilePath, Redemption.RedemptionSaveAsType.olRFC822);
                    //File.OpenRead(sFilePath);
                    StreamReader s = new StreamReader(sFilePath);

                    completeMail = (s.ReadToEnd() + "\r\n" + pobjProfile.ReportEndText);
                    s.Close();

                    // cleanup and delete the file
                    File.Delete(sFilePath);


                    // -------------------------------
                    // Below part is replaced by the above part... due to problems with retrieving MIME headers in outlook.
                    // -------------------------------


                    //string msgHeaders = (string)safeMail.get_Fields((int)X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS);

                    //string msgOrigSubject = pobjMessage.Subject;

                    ////' Test to see if message is HTML.
                    //string msgBody = (string)safeMail.get_Fields((int)X4UMapi.PR_BODY_HTML);
                    ////' This sometimes returns nothing for an HTML email, so
                    ////' we'll do a quick check and use HTMLBody if it is empty
                    //if (msgBody == "") msgBody = safeMail.HTMLBody;
                    //if (msgBody != null)
                    //{
                    //    if (msgBody == "" || msgBody.Contains("<!-- Converted from text/plain format -->"))
                    //    {
                    //        msgBody = (string)safeMail.get_Fields((int)X4UMapi.PR_BODY);
                    //        msgIsHTML = false;
                    //    }
                    //    else
                    //    {
                    //        msgIsHTML = true;
                    //    }
                    //}
                    //if (pobjProfile.CleanHeaders)
                    //    msgHeaders = RepairHeaders(msgHeaders, msgIsHTML);

                    //completeMail = (msgHeaders + msgBody + "\r\n" + pobjProfile.ReportEndText);

                }
            }
            else
            {
                string msgHeaders = (string)safeMail.get_Fields((int)X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS);
                string msgOrigSubject = string.Empty;
                if (pobjMessage is Outlook.MailItem)
                   msgOrigSubject = ((Outlook.MailItem)pobjMessage).Subject;
                else if (pobjMessage is Outlook.PostItem)
                   msgOrigSubject = ((Outlook.PostItem)pobjMessage).Subject;

                //' Test to see if message is HTML.
                string msgBody = (string)safeMail.get_Fields((int)X4UMapi.PR_BODY_HTML);
                //' This sometimes returns nothing for an HTML email, so
                //' we'll do a quick check and use HTMLBody if it is empty
                if (msgBody == "") msgBody = safeMail.HTMLBody;
                if (msgBody != null)
                {
                    if (msgBody == "" || msgBody.Contains("<!-- Converted from text/plain format -->"))
                    {
                        msgBody = (string)safeMail.get_Fields((int)X4UMapi.PR_BODY);
                        msgIsHTML = false;
                    }
                    else
                    {
                        msgIsHTML = true;
                    }
                }
                if (pobjProfile.CleanHeaders)
                    msgHeaders = RepairHeaders(msgHeaders, msgIsHTML);

                completeMail = (msgHeaders + msgBody + "\r\n" + pobjProfile.ReportEndText);
            }


            return completeMail;
        }

        /// <summary>
        /// This function splits the mail and reasembles it as the correct mime mail...
        /// I advice against not using the Redemption library as it is extreamly hard to
        /// recreate the functionallity used here without.
        /// The base problem is to retrieve MIME properties larger than 32/64 KB. this requieres 
        /// the use of the MIME::OpenProperty which I find to be nearly impossible to use from C#.
        /// </summary>
        /// <param name="Message">Redemtion SafemailItem with the outlook message as item.</param>
        /// <returns>The Message as mime text</returns>
        private static string GetMimeMail(SafeMailItem pobjMessage)
        {
            string messageBody = string.Empty;
            string HTMLBody = string.Empty;
            string transportHeader = string.Empty;
            string cleanHeader = string.Empty;
            object obj = Activator.CreateInstance(Type.GetTypeFromProgID("safemail.safemailMailItem"));
            Redemption.SafeMailItem safeMessage = (Redemption.SafeMailItem)obj;//new SafeMailItem();
            //SafeMailItem safeMessage = new SafeMailItem();
            safeMessage.Item = pobjMessage;
            List<string> subHeaders;
            int attachmentCounter = 0;

            List<string> attachmentstrings = new List<string>();

            /* // this part is replaced by use of the redemption lib below.
            transportHeader = X4UMapi.GetMessageProperty(Message.MAPIOBJECT, X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS);
            HTMLBody = X4UMapi.GetMessageProperty(Message.MAPIOBJECT, X4UMapi.PR_BODY_HTML);
            messageBody = X4UMapi.GetMessageProperty(Message.MAPIOBJECT, X4UMapi.PR_BODY);

            if (HTMLBody == String.Empty)
            {
                HTMLBody = Message.HTMLBody;
            }
            if (messageBody == String.Empty)
            {
                messageBody = Message.Body;
            }
            */
            // capture the base information from the email
            transportHeader = (string)safeMessage.get_Fields((int)X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS);
            if (transportHeader == null) transportHeader = string.Empty;
            HTMLBody = (string)safeMessage.get_Fields((int)X4UMapi.PR_BODY_HTML);
            if (HTMLBody == null) HTMLBody = string.Empty;

            messageBody = (string)safeMessage.get_Fields((int)X4UMapi.PR_BODY);
            if (messageBody == null) messageBody = string.Empty;


            // Get the transportheaders if attachments is messages/emails.
            string mainheader = "";
            subHeaders = GetSubmessageHeaders(transportHeader, out mainheader);

            if (safeMessage.Attachments.Count > 0)
            {
                try
                {
                    // loop through the attachments
                    Attachments atts = safeMessage.Attachments;
                    foreach (object oAtt in atts)
                    {
                        string attachmentname;
                        string themessage;

                        byte[] attachmentRawData;
                        object objData = new object();

                        BinaryReader br;
                        FileStream theFile;


                        Attachment theAttachment = oAtt as Attachment;

                        //If the attachment is an embedded item (message) then strip and process the message.
                        if (theAttachment.Type == 5)//olEmbeddeditem
                        {
                            Redemption.SafeMailItem embeddedMessage = (Redemption.SafeMailItem)Activator.CreateInstance(Type.GetTypeFromProgID("safemail.safemailMailItem"));
                            //SafeMailItem embeddedMessage = new SafeMailItem();

                            embeddedMessage.Item = theAttachment.EmbeddedMsg;
                            embeddedMessage.set_Fields((int)X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS, (object)subHeaders[attachmentCounter]);
                            // recursive :-(

                            themessage = GetMimeMail(embeddedMessage);
                            //using (StreamWriter st = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + theAttachment.FileName))
                            //{
                            //    st.Write(themessage);
                            //    st.Close();
                            //}
                            attachmentstrings.Add(themessage);
                            attachmentCounter++;

                        }
                        else// the attachment is a regular attachment.. 
                        {
                            // Save the attachment to disk and read the file as a binary stream for later Base64 encoding
                            attachmentname = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + theAttachment.FileName;
                            if (!File.Exists(attachmentname))
                            {

                                theAttachment.SaveAsFile(attachmentname);
                                theFile = File.Open(attachmentname, FileMode.Open);
                                br = new BinaryReader(theFile);
                                attachmentRawData = new byte[theFile.Length];
                                br.Read(attachmentRawData, 0, (int)theFile.Length);
                                br.Close();
                                File.Delete(attachmentname);
                            }
                            else// just in case if name clash
                            {
                                theAttachment.SaveAsFile(attachmentname + "_");

                                theFile = File.Open(attachmentname + "_", FileMode.Open);
                                br = new BinaryReader(theFile);
                                attachmentRawData = new byte[theFile.Length];
                                br.Read(attachmentRawData, 0, (int)theFile.Length);
                                br.Close();
                                File.Delete(attachmentname + "_");

                            }
                            // re
                            string theRawString = SpamGrabberCommon.Base64.Encode(attachmentRawData);
                            List<string> lstSplitStrings = new List<string>();
                            int lPos = 0;
                            while (lPos < theRawString.Length)
                            {
                                if (theRawString.Length - lPos > 73)
                                {
                                    lstSplitStrings.Add(theRawString.Substring(lPos, 73));
                                    lPos += 73;
                                }
                                else
                                {
                                    lstSplitStrings.Add(theRawString.Substring(lPos, theRawString.Length - lPos));
                                    lPos = theRawString.Length;
                                }

                            }
                            theRawString = string.Join("\r\n", lstSplitStrings.ToArray());
                            attachmentstrings.Add(theRawString);
                        }

                    }
                }
                catch (System.Exception e) // catch all exception for silent falure
                {
                    // catch all exceptions
                }


            }
            // return the entire mail
            if (subHeaders == null || subHeaders.Count == 0)
            {// run mimeconvertion with the original header
                return ConvertMime(transportHeader, messageBody, HTMLBody, attachmentstrings.ToArray());
            }
            else
            { // run mime conversion with the stripped main header
                return ConvertMime(mainheader, messageBody, HTMLBody, attachmentstrings.ToArray());
            }

        }

        /// <summary>
        /// This function processes a transportheader string and splits it to
        /// a list if sub emails and a a string containing the parent part of the 
        /// original mail transport header.
        /// </summary>
        /// <param name="header">The transport header to process</param>
        /// <param name="mainheader">out the part of the original header that does not
        /// relate to attached emails.</param>
        /// <returns>A list of sum email headers</returns>
        private static List<string> GetSubmessageHeaders(string pstrHeader, out string pstrMainheader)
        {
            List<string> lstHeader = new List<string>();
            List<string> lstCombined = new List<string>();
            List<string> lstSubmails = new List<string>();
            List<string> lstMainheader = new List<string>();
            int linecounter = 0;
            string boundaryIdentifier = string.Empty;
            string boundaryIdentifierParent = string.Empty;

            if (pstrHeader == null)
            {
                pstrMainheader = "";
                return null;
            }

            string[] del = new string[] { "\r\n" };
            lstHeader.AddRange(pstrHeader.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None));
            lstHeader[0] = lstHeader[0].Replace("Microsoft Mail Internet Headers Version 2.0\r\n", "");

            Dictionary<string, string> lstBoundaries = GetBoundaryIdTree(lstHeader);

            // traverse through the headers
            foreach (string boundary in lstHeader)
            {
                linecounter++;
                if (boundary.StartsWith("Received: ") && linecounter > 1)// start of new mail not including the first mail header(parent)
                {
                    lstCombined.Clear();
                    boundaryIdentifier = GetBoundaryId(boundary);
                    lstCombined.Add(boundary);
                    if (boundaryIdentifier == "")
                    {
                        lstSubmails.Add(string.Join("\r\n\r\n", lstCombined.ToArray()));// remember to keep the CRLFs
                    }
                }
                else if (boundary.Contains(boundaryIdentifier) && boundaryIdentifier != "")
                {
                    lstCombined.Add(boundary);
                    if (boundary.Contains(boundaryIdentifier + "--"))// end if boundary
                    {
                        lstBoundaries.TryGetValue(boundaryIdentifier, out boundaryIdentifierParent);
                        if (boundaryIdentifierParent == "")// end of message
                        {
                            lstSubmails.Add(string.Join("\r\n\r\n", lstCombined.ToArray()));// remember to keep the CRLFs
                        }
                        else //(boundaryIdentifierParent != "") // just end of boundary
                        {
                            boundaryIdentifier = boundaryIdentifierParent;

                        }

                    }
                    else // not end of boundary
                    {
                        string bi = GetBoundaryId(boundary);
                        if (bi != "")//process a new boundary part
                        {

                            boundaryIdentifier = bi;
                        }
                    }
                }
                else
                {
                    lstMainheader.Add(boundary);
                }
            }
            pstrMainheader = string.Join("\r\n\r\n", lstMainheader.ToArray()); // create the out string

            return lstSubmails; // return the subheaders
        }

        /// <summary>
        /// the original VB code for repairing headers. This is used if selected by user.
        /// </summary>
        /// <param name="header">The header to fix</param>
        /// <param name="msgIsHTML">Wheter or not it is a HTML mail</param>
        /// <returns>The fixed header</returns>
        private static string RepairHeaders(string pstrHeader, bool pblnMsgIsHTML)
        {
            List<string> lstHeader = new List<string>();
            List<string> tempLines = new List<string>();
            string removeString, removeString2;
            bool headerFlag;


            string[] del = new string[] { "\r\n" };
            lstHeader.AddRange(pstrHeader.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None));

            string msgHeaders = "";
            string tempHeader = "";
            //' OK, first we need to break off any boundary headers which are
            //' caused by attachments etc. These are easy to get rid of, as
            //' they always appear two CrLfs after the main headers, so
            //' we just do a split on a double CrLf and take the first part
            msgHeaders = lstHeader[0];

            //' Loop through each line in the headers and remove any
            //' Content-type lines or Microsoft headers
            removeString = "Content-Type";
            removeString2 = "Microsoft Mail Internet Headers Version 2.0";
            headerFlag = false;
            tempLines.AddRange(msgHeaders.Split(del, StringSplitOptions.None));

            foreach (string tempLine in tempLines)
            {
                if (tempLine != string.Empty)
                {
                    if (headerFlag)
                    {
                        //' Flag is on, so we need to see if this is a normal header or
                        //' a continuation header
                        if (tempLine.Substring(0, 1) == " " || tempLine.Substring(0, 1) == "\t")
                        { }   // ' It is a continuation, so do nothing
                        else
                        {
                            //' New header, check to make sure it isn't the
                            //' MS one, and add it
                            if (tempLine.Contains(removeString2))
                            {    //' Just add the header
                                tempHeader += tempLine + "\r\n";
                            }
                            //' Reset the flag
                            headerFlag = false;
                        }
                    }
                    else
                    {  //' Flag is off, so check to see if this is a content type header
                        if (tempLine.Contains(removeString))
                        {    //' This is a content type header, so set the flag and ignore the header
                            headerFlag = true;
                        }
                        else
                        {
                            //' Not a content type header, so check for the MS header and add
                            if (!tempLine.Contains(removeString2))
                            {
                                //' Just add the header
                                tempHeader += tempLine + "\r\n";
                            }
                        }
                    }
                }
            }

            //' Add the correct content type header at the end
            tempHeader = tempHeader + "Content-Type: text/";
            if (pblnMsgIsHTML)
            {
                tempHeader = tempHeader + "html;";
            }
            else
            {
                tempHeader = tempHeader + "plain;";
            }

            //' Set the reporting headers to the cleaned headers
            return tempHeader + "\r\n\r\n";


        }

        /// <summary>
        /// Test if the message has a valid header for MIME operations...
        /// Faulty headers should be captured here.
        /// </summary>
        /// <param name="pobjMessage"></param>
        /// <returns></returns>
        private static bool IsValidMIMEHeader(SafeMailItem pobjMessage)
        {
            List<string> lstHeader = new List<string>();
            string msgHeaders = string.Empty;

            string strBoundary = string.Empty;
            bool blnOK = false;

            if (pobjMessage != null)// check if we actually got an object.
                msgHeaders = (string)pobjMessage.get_Fields((int)X4UMapi.PR_TRANSPORT_MESSAGE_HEADERS);
            else
                return blnOK; // escape the function with no valid mime header

            if (String.IsNullOrEmpty(msgHeaders))
                return blnOK;// escape the function with no valid mime header

            lstHeader.AddRange(msgHeaders.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None));

            if (lstHeader.Count <= 0)
                return blnOK;// escape the function with no valid mime header

            strBoundary = GetBoundaryId(lstHeader[0]);
            if (strBoundary != "")
            {
                foreach (string line in lstHeader)
                {
                    if (line.Contains(strBoundary + "--"))
                        blnOK = true;
                }
                return blnOK;
            }


            return true;
        }
    }


}
