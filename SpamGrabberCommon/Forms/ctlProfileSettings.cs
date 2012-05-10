using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using SpamGrabberCommon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Reflection;

namespace SpamGrabberControl
{
    public partial class ctlProfileSettings : UserControl
    {
        #region Class Data

        private Profile _objProfile;

        #endregion

        #region Constructor

        public ctlProfileSettings()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties

        /// <summary>
        /// The profile to use for this instance of the settings
        /// </summary>
        public Profile Profile
        {
            get
            {
                return this._objProfile;
            }
            set
            {
                this._objProfile = value;
                if (value != null)
                {
                    LoadSettings();
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Loads the settings from the selected profile object
        /// </summary>
        private void LoadSettings()
        {
            // Report address tab
            this.lstToAddresses.Items.Clear();
            foreach (string strAddress in this._objProfile.ToAddresses)
            {
                if (string.IsNullOrEmpty(strAddress) == false)
                {
                    lstToAddresses.Items.Add(strAddress);
                }
            }
            this.lstBccAddresses.Items.Clear();
            foreach (string strAddress in this._objProfile.BccAddresses)
            {
                if (string.IsNullOrEmpty(strAddress) == false)
                {
                    lstBccAddresses.Items.Add(strAddress);
                }
            }
            this.chkDeleteAfterReport.Checked = this._objProfile.DeleteAfterReport;
            this.chkMarkAsRead.Checked = this._objProfile.MarkAsReadAfterReport;
            this.chkMoveToFolder.Checked = this._objProfile.MoveToFolderAfterReport;

            this.txtMoveToFolder.Text = GetFolderById(this._objProfile.MoveFolderName,this._objProfile.MoveFolderStoreId);

            // Report text tab
            this.txtReportSubject.Text = this._objProfile.ReportSubject;
            this.txtReportEndText.Text = this._objProfile.ReportEndText;
            this.txtMessageBody.Text = this._objProfile.MessageBody;

            // Report options tab
            chkAskVerify.Checked = this._objProfile.AskVerify;
            chkKeepCopy.Checked = this._objProfile.KeepCopy;
            chkSendMultiple.Checked = this._objProfile.SendMultiple;
            chkUseRfc.Checked = this._objProfile.UseRFC;
            chkSendReceive.Checked = this._objProfile.SendReceiveAfterReport;
            chkCleanHeaders.Checked = this._objProfile.CleanHeaders;
        }

        private void SyncToAddresses()
        {
            // Clear the collection
            this._objProfile.ToAddresses.Clear();

            // Repopulate with the current addresses
            foreach (string strAddress in lstToAddresses.Items)
            {
                this._objProfile.ToAddresses.Add(strAddress);
            }
        }

        private void SyncBccAddresses()
        {
            // Clear the collection
            this._objProfile.BccAddresses.Clear();

            // Repopulate with the current addresses
            foreach (string strAddress in lstBccAddresses.Items)
            {
                this._objProfile.BccAddresses.Add(strAddress);
            }
        }

        #endregion

        #region Event Handlers

        private void btnDelTo_Click(object sender, EventArgs e)
        {
            lstToAddresses.Items.Remove(lstToAddresses.SelectedItem);
            SyncToAddresses();
        }

        private void btnAddTo_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtAddTo.Text) == false)
            {
                // Make sure the email is valid
                if (SGGlobals.IsEmailValid(txtAddTo.Text) == false)
                {
                    MessageBox.Show("Email address does not appear to be valid!", "Error creating profile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                lstToAddresses.Items.Add(txtAddTo.Text);
                txtAddTo.Text = "";
            }
            SyncToAddresses();
        }

        private void btnAddBcc_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtAddBcc.Text) == false)
            {
                // Make sure the email is valid
                if (SGGlobals.IsEmailValid(txtAddBcc.Text) == false)
                {
                    MessageBox.Show("Email address does not appear to be valid!", "Error creating profile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                lstBccAddresses.Items.Add(txtAddBcc.Text);
                txtAddBcc.Text = "";
            }
            SyncBccAddresses();
        }

        private void btnDelBcc_Click(object sender, EventArgs e)
        {
            lstBccAddresses.Items.Remove(lstBccAddresses.SelectedItem);
            SyncBccAddresses();
        }
       
        private void btnFolderSelect_Click(object sender, EventArgs e)
        {
            object missing = Missing.Value;

            //get the Outlook Application Object
            Outlook.Application outlookApplication = new Outlook.Application();

            //get the namespace object
            Outlook.NameSpace nameSpace = outlookApplication.GetNamespace("MAPI");

            //Logon to Session, here we use an already opened Outlook session
            nameSpace.Logon(missing, missing, false, false);

            //get the InboxFolder
            Outlook.MAPIFolder Folder = nameSpace.PickFolder();
            if (Folder != null)
            {
                txtMoveToFolder.Text = Folder.Name;
                this.Profile.MoveFolderName = Folder.EntryID;
                this.Profile.MoveFolderStoreId = Folder.StoreID;
            }

            //release used resources
            if (Folder != null)
                Marshal.ReleaseComObject(Folder);

            //logof from namespace
            nameSpace.Logoff();

            //release resources
            Marshal.ReleaseComObject(nameSpace);
            Marshal.ReleaseComObject(outlookApplication.Application);
        }

        private void ctlProfileSettings_Load(object sender, EventArgs e)
        {
            ToolTip ttProfileSettings = new ToolTip();

            // Set up the delays for the ToolTip.
            ttProfileSettings.AutoPopDelay = 5000;
            ttProfileSettings.InitialDelay = 1000;
            ttProfileSettings.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            ttProfileSettings.ShowAlways = true;

            // Set the tooltips for the control
            ttProfileSettings.SetToolTip(chkAskVerify, "Asks for confermation every time you try to report a message.");
            ttProfileSettings.SetToolTip(chkCleanHeaders, "Tells SpamGrabber to process and remove \r\nany message headers it does not understand.");
            ttProfileSettings.SetToolTip(chkDeleteAfterReport, "Tells SpamGrabber to delete the selected \r\nmessages after they have beed reportet.");
            ttProfileSettings.SetToolTip(chkKeepCopy, "Tells SpamGrabber to keep a copy of the report.");
            ttProfileSettings.SetToolTip(chkMarkAsRead, "Tells SpamGrabber to mark the selected \r\nmessages as read after the report has been send.");
            ttProfileSettings.SetToolTip(chkMoveToFolder, "Tells SpamGrabber to move the \r\nselected messages to the designated folder.");
            ttProfileSettings.SetToolTip(chkSendMultiple, "Tells SpamGrabber to send all \r\nthe selected messages as one report \r\ninstead of one report pr. message.");
            ttProfileSettings.SetToolTip(chkSendReceive, "Tells SpamGrabber to force a Send and Recieve \r\noperation after the report has been submitted.");
            ttProfileSettings.SetToolTip(chkUseRfc, "Tells SpamGrabber to use the RFC822 message \r\nstandard when reporting the message.");

        }
        #endregion

        #region Synchronisation methods

        private void chkDeleteAfterReport_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.DeleteAfterReport = chkDeleteAfterReport.Checked;
            if (chkDeleteAfterReport.Checked == true)
            {
                chkMarkAsRead.Checked = false;
                chkMarkAsRead.Enabled = false;

                chkMoveToFolder.Checked = false;
                chkMoveToFolder.Enabled = false;

                btnFolderSelect.Enabled = false;
                txtMoveToFolder.Enabled = false;
            }
            else 
            {
                
                chkMarkAsRead.Enabled = true;

                chkMoveToFolder.Enabled = true;

                btnFolderSelect.Enabled = true;
                txtMoveToFolder.Enabled = true;
            }
        }

        private void chkMarkAsRead_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.MarkAsReadAfterReport = chkMarkAsRead.Checked;
        }

        private void chkMoveToFolder_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.MoveToFolderAfterReport = chkMoveToFolder.Checked;
            this.txtMoveToFolder.Enabled = chkMoveToFolder.Checked;
        }

        private void txtReportSubject_TextChanged(object sender, EventArgs e)
        {
            this._objProfile.ReportSubject = txtReportSubject.Text;
        }

        private void txtReportEndText_TextChanged(object sender, EventArgs e)
        {
            this._objProfile.ReportEndText = txtReportEndText.Text;
        }

        private void txtMessageBody_TextChanged(object sender, EventArgs e)
        {
            this._objProfile.MessageBody = txtMessageBody.Text;
        }

        private void chkAskVerify_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.AskVerify = chkAskVerify.Checked;
        }

        private void chkKeepCopy_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.KeepCopy = chkKeepCopy.Checked;
        }

        private void chkSendMultiple_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.SendMultiple = chkSendMultiple.Checked;
        }

        private void chkUseRfc_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.UseRFC = chkUseRfc.Checked;
        }

        private void chkSendReceive_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.SendReceiveAfterReport = chkSendReceive.Checked;
        }

        private void chkCleanHeaders_CheckedChanged(object sender, EventArgs e)
        {
            this._objProfile.CleanHeaders = chkCleanHeaders.Checked;
        }

        #endregion
        
        /// <summary>
        /// Get the foldername from the FolderId and StoreId
        /// </summary>
        /// <param name="FolderId"></param>
        /// <param name="StoreId"></param>
        /// <returns></returns>
        private string GetFolderById(string FolderId, string StoreId)
        {
            object missing = Missing.Value;
            string foldername = string.Empty;

            // if invalid input data then exit returning empty string
            if (FolderId == string.Empty || StoreId == string.Empty)
                return string.Empty;

            //get the Outlook Application Object
            Outlook.Application outlookApplication = new Outlook.Application();

            //get the namespace object
            Outlook.NameSpace nameSpace = outlookApplication.GetNamespace("MAPI");

            //Logon to Session, here we use an already opened Outlook session
            nameSpace.Logon(missing, missing, false, false);

            //get the InboxFolder
            Outlook.MAPIFolder Folder = nameSpace.GetFolderFromID(FolderId,StoreId);
            if (Folder != null)
            {
              foldername = Folder.Name;
               
            }

            //release used resources

            Marshal.ReleaseComObject(Folder);

            //logof from namespace
            nameSpace.Logoff();

            //release resources
            Marshal.ReleaseComObject(nameSpace);
            Marshal.ReleaseComObject(outlookApplication.Application);

            return foldername;
        }
 
    }
}