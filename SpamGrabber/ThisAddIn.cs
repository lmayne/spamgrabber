using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using SpamGrabberControl;
using SpamGrabberCommon;
using X4U.Outlook;
using System.IO;
using Microsoft.Office.Core;

namespace SpamGrabber2003
{
    public partial class ThisAddIn
    {
        #region Class Data

        private Outlook.Explorer _objExplorer = null;
        private Office.CommandBarButton _cbbDefaultHam;
        private Office.CommandBarButton _cbbDefaultSpam;
        private Office.CommandBarButton _cbbReportToProfile;
        private Office.CommandBarButton _cbbCopyToClipboard;
        private Office.CommandBarButton _cbbSendToSupport;
        private Office.CommandBarButton _cbbPreview;
        private Office.CommandBarButton _cbbOptions;
        private Office.CommandBar _cbSpamGrabber;
        private Office.CommandBars _cbcCommandBars;
        //private Outlook.Application _appApplication;
        //private Outlook.Explorers _objExplorers;


        #endregion

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            _objExplorer = this.Application.ActiveExplorer();

            // Register event so we handle when a user switch to a different folder
            _objExplorer.FolderSwitch += new
                Outlook.ExplorerEvents_10_FolderSwitchEventHandler(
                    Explore_FolderSwitch);

            // Register event so we handle when a user switch to a different item
            _objExplorer.SelectionChange += new
                Outlook.ExplorerEvents_10_SelectionChangeEventHandler(
                    Explorer_SelectionChange);

            // handle multible explorers.
            if (this.Application.Explorers.Count > 0)
            {
                // This event helps keep track of changes to the toolbars in the application.
                _cbcCommandBars = _objExplorer.CommandBars;
                _cbcCommandBars.OnUpdate += new _CommandBarsEvents_OnUpdateEventHandler(CommandBars_OnUpdate);

                // Pupulate menu
                BuildMenu();
            }
        }
        

        private void CommandBars_OnUpdate()
        {
            // highly active event... update position data only....
            GlobalSettings.CommandBarTop = _cbSpamGrabber.Top;
            GlobalSettings.CommandBarLeft = _cbSpamGrabber.Left;
            GlobalSettings.CommandBarVisible = _cbSpamGrabber.Visible;
            GlobalSettings.CommandBarPosition = (int)_cbSpamGrabber.Position;
            GlobalSettings.CommandBarRowIndex = _cbSpamGrabber.RowIndex;
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            SpamGrabberCommon.CommandBars.SaveCommandBarSettings(_cbSpamGrabber);

        }

        #region Interface
        /// <summary>
        /// Enables our menubar on folders that contain mail items
        /// </summary>
        private void Explore_FolderSwitch()
        {
            String folderWebView;
            folderWebView = _objExplorer.CurrentFolder.WebViewURL;

            if (folderWebView == null)
            {
                if (_objExplorer.CurrentFolder.DefaultItemType == Outlook.OlItemType.olMailItem || _objExplorer.CurrentFolder.DefaultItemType == Outlook.OlItemType.olPostItem)
                {
                    //_cbSpamGrabber.Enabled = true;
                    GrayoutAllButtonsNot();
                }
                else
                {
                    //_cbSpamGrabber.Enabled = false;
                    GrayoutAllButtons();
                }
            }
            else
            {
                if (!folderWebView.Contains("outlook.htm"))
                {

                    if (_objExplorer.CurrentFolder.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        //_cbSpamGrabber.Enabled = true;
                        GrayoutAllButtonsNot();

                    }
                    else
                    {
                        //_cbSpamGrabber.Enabled = false;
                        GrayoutAllButtons();
                    }
                }
                else
                {
                    //_cbSpamGrabber.Enabled = false;
                    GrayoutAllButtons();
                }
            }

        }

        /// <summary>
        /// Checks to ensure we can handle the selected change in selection
        /// </summary>
        private void Explorer_SelectionChange()
        {
            //bool blnHasInvalidItem = false;
            //int explorerCount = .Explorers.Count;
            if (_objExplorer.Selection.Count > 0)
            {

                if (_objExplorer.Selection.Count == 1)// only one item in list
                {

                    foreach (object objItem in _objExplorer.Selection)
                    {
                        if (objItem is Outlook.MailItem || objItem is Outlook.PostItem)
                        {

                            if (!GlobalSettings.UseStandardToolbar)
                            {
                                // We have a mail item
                                //_cbSpamGrabber.Enabled = true;
                                GrayoutAllButtonsNot();
                            }
                            else
                            {
                                ShowHideButtons();
                            }
                        }
                        else
                        {
                            if (!GlobalSettings.UseStandardToolbar)
                            {
                                //_cbSpamGrabber.Enabled = false;
                                GrayoutAllButtons();
                            }
                            else
                            {
                                // hide the spamgrabber buttons.
                                HideAllButtons();
                            }

                        }
                    }
                }
                else
                {
                    if (!GlobalSettings.UseStandardToolbar)
                    {
                        // We have a mail item
                        //_cbSpamGrabber.Enabled = true;
                        GrayoutAllButtonsNot();
                    }
                    else
                    {
                        ShowHideButtons();
                    }
                }
            }
            else
            {
                if (!GlobalSettings.UseStandardToolbar)
                {
                    //_cbSpamGrabber.Enabled = false;
                    GrayoutAllButtons();
                }
                else
                {
                    // hide the spamgrabber buttons.
                    HideAllButtons();
                }
            }
        }
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

        }

        #endregion

        #region Methods

        /// <summary>
        /// Builds the menu, adds the buttons, and wires up the button events
        /// </summary>
        private void BuildMenu()
        {
            int basePosition = 0;// start position for buttons

            try
            {
                // Destroy old commandbar - if any -
                foreach (Office.CommandBar bar in _objExplorer.CommandBars)
                {
                    if (bar.Name == GlobalSettings.COMMANDBAR_NAME)
                    {
                        bar.Delete();
                        break;
                    }
                }


                if (GlobalSettings.UseStandardToolbar)
                {
                    foreach (Office.CommandBar bar in _objExplorer.CommandBars)
                    {
                        if (bar.Name == "Standard") // use standard toolbar
                            _cbSpamGrabber = bar;
                    }
                    basePosition = _cbSpamGrabber.Controls.Count;
                }
                else
                {
                    // Start with a fresh one...
                    _cbSpamGrabber = _objExplorer.CommandBars.Add(GlobalSettings.COMMANDBAR_NAME,
                        Office.MsoBarPosition.msoBarTop, false, true);
                    basePosition = 0;

                }


                // Do registry work
                try
                {
                    if (SGGlobals.GetUserConfigurationKey(SGGlobals.GetBaseConfigurationKey()) == null)
                    {
                        MessageBox.Show(@"This appears to be the first time you have run
                            SpamGrabber. Please remember to set a default report address in
                            the options page");

                        SGGlobals.CreateMachineConfigurationKey(SGGlobals.GetBaseConfigurationKey());
                    }
                }
                catch (System.Exception ex)
                {
                    string error = String.Format(
                        "Unable to check registry values: {0}", ex.ToString());

                    MessageBox.Show(error);
                }

                /****************
                 Create buttons    
                ****************/
                // Declare helper boolean
                bool blnIsVisible = false;

                // Spam button
                //_cbbDefaultSpam = CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                //    "Report to Spam profile(s)", "Report to Spam profile(s)", "Report to Spam profile(s)",
                //    352, Office.MsoButtonStyle.msoButtonIcon,
                //    true, true, basePosition + 1);
                _cbbDefaultSpam = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                    "Report Spam", "Report to Spam profile(s)", "Report to Spam profile(s)",
                    "spamgrab_red", "spamgrab_green_mask", Office.MsoButtonStyle.msoButtonIconAndCaption,
                    true, true, basePosition + 1);
                // Spam event handler
                _cbbDefaultSpam.Click += new
                    Office._CommandBarButtonEvents_ClickEventHandler(Spam_Click);

                // Ham button
                blnIsVisible = GlobalSettings.ShowHamButton;
                //_cbbDefaultHam = CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                //    "Report to Ham profile(s)", "Report to Ham profile(s)", "Report to Ham profile(s)",
                //    351, Office.MsoButtonStyle.msoButtonIcon,
                //    false, blnIsVisible, basePosition + 2);
                _cbbDefaultHam = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                    "Report Ham", "Report to Ham profile(s)", "Report to Ham profile(s)",
                    "spamgrab_green", "spamgrab_green_mask", Office.MsoButtonStyle.msoButtonIconAndCaption,
                    false, blnIsVisible, basePosition + 2);
                // Ham event handler
                _cbbDefaultHam.Click += new
                    Office._CommandBarButtonEvents_ClickEventHandler(Ham_Click);

                // Report to profile button  
                blnIsVisible = GlobalSettings.ShowSelectButton;
                //_cbbReportToProfile = CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                //     "Report to profile...", "Report to profile...", "Report to profile...",
                //    341, Office.MsoButtonStyle.msoButtonIcon,
                //    false, blnIsVisible, basePosition + 3);
                _cbbReportToProfile = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                     "Report to", "Report to profile...", "Report to profile...",
                    "spamgrabber_RG", "spamgrab_green_mask", Office.MsoButtonStyle.msoButtonIconAndCaption,
                    false, blnIsVisible, basePosition + 3);
                // report to profile event handler  
                _cbbReportToProfile.Click += new
                   Office._CommandBarButtonEvents_ClickEventHandler(SelectSpam_Click);


                // Copy button
                blnIsVisible = GlobalSettings.ShowCopyButton;
                //_cbbCopyToClipboard = CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                //    "Copy to clipboard", "Copy to clipboard", "Copy to clipboard",
                //    19, Office.MsoButtonStyle.msoButtonIcon,
                //    true, blnIsVisible, basePosition + 4);
                _cbbCopyToClipboard = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                    "Copy to clipboard", "Copy to clipboard", "Copy to clipboard",
                    "spamgrab_copy", "spamgrab_copy_mask", Office.MsoButtonStyle.msoButtonIcon,
                    true, blnIsVisible, basePosition + 4);
                // Copy event handler
                _cbbCopyToClipboard.Click += new
                    Office._CommandBarButtonEvents_ClickEventHandler(Copy_Click);

                // Support button
                blnIsVisible = GlobalSettings.ShowSupportButton;
                //_cbbCopyToClipboard = CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                //    "Copy to clipboard", "Copy to clipboard", "Copy to clipboard",
                //    19, Office.MsoButtonStyle.msoButtonIcon,
                //    true, blnIsVisible, basePosition + 4);
                _cbbSendToSupport = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                    "Send to support", "Send to support", "Send to support",
                    "spamgrab_support", "spamgrab_support_mask", Office.MsoButtonStyle.msoButtonIcon,
                    false, blnIsVisible, basePosition + 5);
                // Copy event handler
                _cbbSendToSupport.Click += new
                    Office._CommandBarButtonEvents_ClickEventHandler(Support_Click);

                // Preview button
                blnIsVisible = GlobalSettings.ShowPreviewButton;
                _cbbPreview = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                    "Safe Preview", "Safe Preview", "Safe Preview",
                    172, Office.MsoButtonStyle.msoButtonIcon,
                    false, blnIsVisible, basePosition + 6);
                // Copy event handler
                _cbbPreview.Click += new
                    Office._CommandBarButtonEvents_ClickEventHandler(Preview_Click);

                // Options button
                blnIsVisible = GlobalSettings.ShowSettingsButton;
                //_cbbOptions = CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                //    "SpamGrabber options", "SpamGrabber options",
                //    "SpamGrabber options", 51, Office.MsoButtonStyle.msoButtonIcon,
                //    true, blnIsVisible, basePosition + 6);
                _cbbOptions = SpamGrabberCommon.CommandBars.CreateCommandBarButton(_cbSpamGrabber,
                    "SpamGrabber options", "SpamGrabber options",
                    "SpamGrabber options", "spamgrab_settings", "spamgrab_settings_mask", Office.MsoButtonStyle.msoButtonIcon,
                    true, blnIsVisible, basePosition + 7);
                //// Options event handler
                _cbbOptions.Click += new
                    Office._CommandBarButtonEvents_ClickEventHandler(Options_Click);



                // Enable menu
                _cbSpamGrabber.Visible = true;
                _cbSpamGrabber.Enabled = true;
                SpamGrabberCommon.CommandBars.LoadCommandBarSettings(_cbSpamGrabber);
                //_cbSpamGrabber.Position = Office.MsoBarPosition.msoBarTop;
            }
            catch (System.Exception ex)
            {
                if (!GlobalSettings.UseStandardToolbar)
                {
                    _cbSpamGrabber.Visible = false;
                    _cbSpamGrabber.Enabled = false;
                }
                else
                {
                    // hide the spamgrabber buttons.
                    HideAllButtons();
                }

                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Grays out all buttons on the commandbar
        /// Function used to simulate disabling the commandbar
        /// </summary>
        private void GrayoutAllButtons()
        {
            _cbbDefaultSpam.Enabled = false;
            _cbbDefaultHam.Enabled = false;
            _cbbReportToProfile.Enabled = false;
            _cbbCopyToClipboard.Enabled = false;
            _cbbSendToSupport.Enabled = false;
            _cbbOptions.Enabled = true;
            _cbbPreview.Enabled = false;
        }
        /// <summary>
        /// Removes the grayout from all buttons
        /// Function used to simulate disabling the commandbar
        /// </summary>
        private void GrayoutAllButtonsNot()
        {
            _cbbDefaultSpam.Enabled = true;
            _cbbDefaultHam.Enabled = true;
            _cbbReportToProfile.Enabled = true;
            _cbbCopyToClipboard.Enabled = true;
            _cbbSendToSupport.Enabled = true;
            _cbbOptions.Enabled = true;
            _cbbPreview.Enabled = true;
        }

        /// <summary>
        /// Hides all the buttons from spamgrabber.
        /// This is used in place of hiding the intire toolbar.
        /// This is added in version 4.0.2
        /// </summary>
        private void HideAllButtons()
        {
            _cbbDefaultSpam.Visible = false;
            _cbbDefaultHam.Visible = false;
            _cbbReportToProfile.Visible = false;
            _cbbCopyToClipboard.Visible = false;
            _cbbSendToSupport.Visible = false;
            _cbbOptions.Visible = false;
            _cbbPreview.Visible = false;
        }

        /// <summary>
        /// Changes the buttons visible state based on global settings
        /// </summary>
        private void ShowHideButtons()
        {
            _cbbDefaultSpam.Visible = true;
            _cbbDefaultHam.Visible = GlobalSettings.ShowHamButton;
            _cbbReportToProfile.Visible = GlobalSettings.ShowSelectButton;
            _cbbCopyToClipboard.Visible = GlobalSettings.ShowCopyButton;
            _cbbSendToSupport.Visible = GlobalSettings.ShowSupportButton;
            _cbbOptions.Visible = GlobalSettings.ShowSettingsButton;
            _cbbPreview.Visible = GlobalSettings.ShowPreviewButton;
        }

        #endregion

        #region Button Event Handlers

        /// <summary>
        /// Reports the selected message(s) to the default Spam profile
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void Spam_Click(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            // Make sure they have set a default
            if (string.IsNullOrEmpty(GlobalSettings.DefaultSpamProfileId))
            {
                MessageBox.Show("You have not selected a default spam reporting profile. " +
                    "Please set one of your profiles to be your default spam profile in the settings page",
                    "No default profile set", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            // process mail for Spam
            ReportMail(SGGlobals.ReportAction.ReportSpam);
        }

        /// <summary>
        /// Reports the selected message(s) to the default Ham profile
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void Ham_Click(Office.CommandBarButton ctrl,
           ref bool cancelDefault)
        {
            if (string.IsNullOrEmpty(GlobalSettings.DefaultHamProfileId))
            {
                MessageBox.Show("You have not selected a default ham reporting profile. " +
                    "Please set one of your profiles to be your default ham profile in the settings page",
                    "No default profile set", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            // process mail for ham
            ReportMail(SGGlobals.ReportAction.ReportHam);
        }

        /// <summary>
        /// Reports the selected message(s) to the selected profile.
        /// It shows a dialog and asks the user to select a profile.
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void SelectSpam_Click(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            if (_objExplorer.Selection.Count > 0)
            {
                // Open the dialog
                frmSelectProfile objForm = new frmSelectProfile();
                if (objForm.ShowDialog() == DialogResult.OK)
                {
                    // process mail
                    ProcessMail(objForm.ReportAction, objForm.SelectedProfile, true);
                }
            }
        }

        /// <summary>
        /// Copies the first selected item in the list to the clipboard
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void Copy_Click(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                Profile objProfile = new Profile(GlobalSettings.DefaultSpamProfileId);
                if (_objExplorer.Selection.Count > 0)
                {
                    List<string> RawMails = new List<string>();

                    foreach (object objItem in _objExplorer.Selection)
                    {
                        if (objItem is Outlook.MailItem || objItem is Outlook.PostItem)
                        {
                            RawMails.Add(MimeConversions.ConvertMail(objItem, objProfile));
                        }
                    }
                    Clipboard.Clear();
                    Clipboard.SetText(string.Join("", RawMails.ToArray()));

                }

            }
            catch (Exception ex) // TODO we should not catch all exceptions
            {
                MessageBox.Show("caught:" + ex.ToString());
            }
        }

        private void Support_Click(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                ReportMail(SGGlobals.ReportAction.ReportSupport);
            }
            catch (Exception ex) // TODO we should not catch all exceptions
            {
                MessageBox.Show("caught:" + ex.ToString());
            }
        }

        /// <summary>
        /// Handles the Preview button (Displays the safe preview form)
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void Preview_Click(Office.CommandBarButton ctrl,
           ref bool cancelDefault)
        {
            if (_objExplorer.Selection.Count > 0)
            {
                frmPreview objPreview = new frmPreview();
                objPreview.ClearItems();
                foreach (object objItem in _objExplorer.Selection)
                {
                    if (objItem is Outlook.MailItem || objItem is Outlook.PostItem)
                        objPreview.Items.Add(objItem);
                }
                objPreview.ShowDialog();
            }
        }

        /// <summary>
        /// Opens the options form
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        private void Options_Click(Office.CommandBarButton ctrl,
           ref bool cancelDefault)
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
            catch (Exception ex) // TODO we should not catch all exceptions
            {
                MessageBox.Show("caught: \r\n" + ex.ToString());
            }
        }

        #endregion

        #region Helper functions
        /// <summary>
        /// This function processes the messages with the different profiles.
        /// If the messages is beeing processed by more than one run, then
        /// the messages is only beeing processed for move, unread and deletion by
        /// the default profile...
        /// </summary>
        /// <param name="theAction">The action to take</param>
        /// <param name="objDefaultProfile">The profile to process the messages with</param>
        /// <param name="isLastProfile">Tells if the messages is being processed for the last time in this run</param>
        private void ProcessMail(SGGlobals.ReportAction pAction, Profile pobjProfile, bool pblnIsLastProfile)
        {
            List<string> attachmentNames = new List<string>();
            //            Profile objDefaultProfile = null;
            bool isSpamReport = false;
            string s = string.Empty;
            try
            {
                if (pAction == SGGlobals.ReportAction.ReportSupport)
                {
                    string to = string.Empty;
                    frmInput input = new SpamGrabberControl.frmInput();
                    if (input.ShowDialog()==DialogResult.Cancel || input.Email==string.Empty)
                    {
                        return;// cancle the operation
                    }
                    // continue the operation
                    to = input.Email;
                    // create the new message and send it to the profile selected
                    Outlook.MailItem newMail;
                    newMail = (Outlook.MailItem)this.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                    s+=  to.Trim() +";";
                    newMail.To = s;

                    newMail.Subject = "Review messages from SpamGrabber " + GlobalSettings.PROG_VERSION;
                    newMail.Body = "Hi support.\nPlease review the attached email(s)";

                    foreach (object selectedItem in _objExplorer.Selection)
                    {
                        Object objItem = selectedItem;

                        if (objItem != null && (objItem is Outlook.MailItem || objItem is Outlook.PostItem))
                        {
                            // send email to support using RFC822
                            //newMail.Attachments.Add(objItem, Outlook.OlAttachmentType.olByValue, missing, missing);
                            string filename = string.Empty;
                            filename = SpamGrabberCommon.MimeConversions.ExportToFile(objItem, pAction, pobjProfile);

                            newMail.Attachments.Add(filename, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, missing, filename);
                            File.Delete(filename);


                            // mark original message as read
                            if (objItem is Outlook.MailItem)
                                ((Outlook.MailItem)objItem).UnRead = false;
                            else if (objItem is Outlook.PostItem)
                                ((Outlook.PostItem)objItem).UnRead= false;

                        }
                        
                    }


                    // Keep copy of report in sendt items?
                    newMail.DeleteAfterSubmit = false;


                    // Send the mail
                    (newMail as Outlook._MailItem).Send();
                    // cleanup temp folder

                }
                else
                {
                    // execute verification if selected.
                    if (pobjProfile.AskVerify)
                    {
                        if (_objExplorer.Selection.Count > 1)
                        {
                            if (MessageBox.Show("Are you sure you want to report these emails?", "SpamGrabber", MessageBoxButtons.YesNo) == DialogResult.No)
                                return;
                        }
                        else
                        {
                            if (MessageBox.Show("Are you sure you want to report this email?", "SpamGrabber", MessageBoxButtons.YesNo) == DialogResult.No)
                                return;
                        }
                    }

                    isSpamReport = pAction == SGGlobals.ReportAction.ReportSpam ? true : false;

                    // process multible items as one report
                    if (pobjProfile.SendMultiple)
                    {
                        // create the new message and send it to the profile selected
                        Outlook.MailItem newMail;
                        newMail = (Outlook.MailItem)this.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                        foreach (string to in pobjProfile.ToAddresses)
                        {
                            if (to != string.Empty)
                            {
                                s += to.Trim() + ";";
                                newMail.To = s;
                            }
                        }

                        foreach (string to in pobjProfile.BccAddresses)
                        {
                            if (to != string.Empty)
                            {
                                s += to.Trim() + ";";
                                newMail.BCC = s;
                            }
                        }
                        newMail.Subject = pobjProfile.ReportSubject;
                        newMail.Body = pobjProfile.MessageBody;

                        foreach (object selectedItem in _objExplorer.Selection)
                        {
                            object objItem = selectedItem;

                            if (objItem != null && (objItem is Outlook.MailItem || objItem is Outlook.PostItem))
                            {

                                string filename = string.Empty;

                                if (pobjProfile.UseRFC)
                                {
                                    newMail.Attachments.Add(objItem, Outlook.OlAttachmentType.olByValue, missing, missing);
                                }
                                else
                                {
                                    filename = SpamGrabberCommon.MimeConversions.ExportToFile(objItem, pAction, pobjProfile);

                                    newMail.Attachments.Add(filename, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, missing, filename);
                                    File.Delete(filename);
                                }
                                if (pblnIsLastProfile)
                                {

                                    // mark original message as read
                                    if (pobjProfile.MarkAsReadAfterReport)
                                    {
                                        if (objItem is Outlook.MailItem)
                                            ((Outlook.MailItem)objItem).UnRead = false;
                                        else if (objItem is Outlook.PostItem)
                                            ((Outlook.PostItem)objItem).UnRead = false;
                                    }
                                    // Delete original message
                                    if (pobjProfile.DeleteAfterReport)// should we hard delete the message
                                    {
                                        // dont delete ham
                                        if (pAction != SGGlobals.ReportAction.ReportHam)
                                        {
                                            if (objItem is Outlook.MailItem)
                                                ((Outlook.MailItem)objItem).Delete();
                                            else if (objItem is Outlook.PostItem)
                                                ((Outlook.PostItem)objItem).Delete();
                                        }
                                    }
                                    // move the message to designated spam folder
                                    else if (pobjProfile.MoveToFolderAfterReport)
                                    {

                                        Outlook.MAPIFolder Folder = this._objExplorer.Application.GetNamespace("MAPI").GetFolderFromID(
                                            pobjProfile.MoveFolderName, pobjProfile.MoveFolderStoreId);

                                        if (objItem is Outlook.MailItem)
                                            ((Outlook.MailItem)objItem).Move(Folder);
                                        else if (objItem is Outlook.PostItem)
                                            ((Outlook.PostItem)objItem).Move(Folder);
                                    }
                                }
                            }
                        }


                        // Keep copy of report in sendt items?
                        if (!pobjProfile.KeepCopy)
                        {
                            newMail.DeleteAfterSubmit = true;
                        }

                        // Send the mail
                        (newMail as Outlook._MailItem).Send();
                        // cleanup temp folder

                    }
                    else // send one report pr. email
                    {

                        foreach (object selectedItem in _objExplorer.Selection)
                        {
                            object objItem = selectedItem;

                            if (objItem != null && (objItem is Outlook.MailItem || objItem is Outlook.PostItem))
                            {

                                string filename = string.Empty;

                                // create the new message and send it to the priofile selected
                                Outlook.MailItem newMail;
                                newMail = (Outlook.MailItem)this.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                                foreach (string to in pobjProfile.ToAddresses)
                                {
                                    if (to != string.Empty)
                                    {
                                        s += to.Trim() + ";";
                                        newMail.To = s;
                                    }
                                }

                                foreach (string to in pobjProfile.BccAddresses)
                                {
                                    if (to != string.Empty)
                                    {
                                        s += to.Trim() + ";";
                                        newMail.BCC = s;
                                    }
                                }
                                newMail.Subject = pobjProfile.ReportSubject;
                                newMail.Body = pobjProfile.MessageBody;


                                if (pobjProfile.UseRFC)
                                {
                                    newMail.Attachments.Add(objItem, Outlook.OlAttachmentType.olByValue, missing, missing);
                                }
                                else
                                {
                                    filename = SpamGrabberCommon.MimeConversions.ExportToFile(objItem, pAction, pobjProfile);

                                    newMail.Attachments.Add(filename, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, missing, filename);
                                    File.Delete(filename);
                                }
                                if (pblnIsLastProfile)
                                {
                                    // mark original message as read
                                    if (pobjProfile.MarkAsReadAfterReport)
                                    {
                                        if (objItem is Outlook.MailItem)
                                            ((Outlook.MailItem)objItem).UnRead = false;
                                        else if (objItem is Outlook.PostItem)
                                            ((Outlook.PostItem)objItem).UnRead = false;
                                    }

                                    // Delete original message
                                    if (pobjProfile.DeleteAfterReport)
                                    {
                                        // dont delete ham
                                        if (pAction != SGGlobals.ReportAction.ReportHam)
                                        {
                                            if (objItem is Outlook.MailItem)
                                                ((Outlook.MailItem)objItem).Delete();
                                            else if (objItem is Outlook.PostItem)
                                                ((Outlook.PostItem)objItem).Delete();
                                        }
                                    }
                                    // move the message to designated spam folder
                                    else if (pobjProfile.MoveToFolderAfterReport)
                                    {

                                        Outlook.MAPIFolder Folder = this._objExplorer.Application.GetNamespace("MAPI").GetFolderFromID(
                                            pobjProfile.MoveFolderName, pobjProfile.MoveFolderStoreId);
                                        if (objItem is Outlook.MailItem)
                                            ((Outlook.MailItem)objItem).Move(Folder);
                                        else if (objItem is Outlook.PostItem)
                                            ((Outlook.PostItem)objItem).Move(Folder);
                                    }
                                }

                                // Keep copy of report in sendt items?
                                if (!pobjProfile.KeepCopy)
                                {
                                    newMail.DeleteAfterSubmit = true;
                                }

                                // Send the mail
                                (newMail as Outlook._MailItem).Send();

                            }

                        }

                    }


                    // force outlook to send recieve
                    if (pobjProfile.SendReceiveAfterReport)
                    {
                        // TODO: Find control ID for the outlook 2007 Send and recieve
                        Office.CommandBarControl btn = this.Application.ActiveExplorer().CommandBars.FindControl(1, 7095, null, null);
                        if (btn != null)
                            btn.Execute();
                    }
                }
            }
            catch (System.Exception ex) // TODO we should not catch all exceptions or display them to the user!
            {
                string message = "Spamgrabber caught this exception: \r\n" + ex.ToString();
                MessageBox.Show(message, "Spamgrabber error");
            }
        }
        /// <summary>
        /// Run the mail report to include running with multiple profiles
        /// </summary>
        /// <param name="theAction"></param>
        private void ReportMail(SGGlobals.ReportAction pAction)
        {
            Profile objDefaultProfile;

            if (pAction == SGGlobals.ReportAction.ReportSpam)
            {

                objDefaultProfile = new Profile(GlobalSettings.DefaultSpamProfileId);

                if (GlobalSettings.ReportToMultipleProfiles == false)
                {
                    ProcessMail(SGGlobals.ReportAction.ReportSpam, objDefaultProfile, true);
                }
                else
                {
                    foreach (Profile theProfile in UserProfiles.ProfileList)
                    {
                        if (theProfile.ProfileType == Profile.eProfileType.SpamProfile || theProfile.ProfileType == Profile.eProfileType.SpamHamProfile)
                        {
                            // don't process the default profile now.
                            if (theProfile.Id != objDefaultProfile.Id)
                                ProcessMail(SGGlobals.ReportAction.ReportSpam, theProfile, false);
                        }
                    }
                    // process the default profile last.
                    ProcessMail(SGGlobals.ReportAction.ReportSpam, objDefaultProfile, true);
                }


            }
            else if (pAction == SGGlobals.ReportAction.ReportHam)
            {
                // Make sure they have set a default


                objDefaultProfile = new Profile(GlobalSettings.DefaultHamProfileId);


                if (GlobalSettings.ReportToMultipleProfiles == false)
                {

                    ProcessMail(SGGlobals.ReportAction.ReportHam, objDefaultProfile, true);
                }
                else
                {
                    foreach (Profile theProfile in UserProfiles.ProfileList)
                    {
                        if (theProfile.ProfileType == Profile.eProfileType.HamProfile || theProfile.ProfileType == Profile.eProfileType.SpamHamProfile)
                        {
                            // don't process the default profile now.
                            if (theProfile.Id != objDefaultProfile.Id)
                                ProcessMail(SGGlobals.ReportAction.ReportHam, theProfile, false);
                        }
                    }
                    // process the default profile last.
                    ProcessMail(SGGlobals.ReportAction.ReportHam, objDefaultProfile, true);
                }

            }
            else if (pAction == SGGlobals.ReportAction.ReportSupport)
            {
                objDefaultProfile = objDefaultProfile = new Profile(GlobalSettings.DefaultSpamProfileId); // Use default spam profile to report to support.
                // process the default profile last.
                ProcessMail(SGGlobals.ReportAction.ReportSupport, objDefaultProfile, true);

            }
        }
        #endregion
    }
}
