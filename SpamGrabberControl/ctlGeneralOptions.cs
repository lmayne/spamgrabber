using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using SpamGrabberCommon;

namespace SpamGrabberControl
{
    public partial class ctlGeneralOptions : UserControl
    {
        public ctlGeneralOptions()
        {
            InitializeComponent();
            UIControl.StopWindowUpdating(this.Handle);
            // Load the initial settings
            this.chkShowPreviewButtonBox.Checked = GlobalSettings.ShowPreviewButton;
            this.chkShowCopyButton.Checked = GlobalSettings.ShowCopyButton;
            this.chkShowHamButton.Checked = GlobalSettings.ShowHamButton;
            this.chkShowSelectButton.Checked = GlobalSettings.ShowSelectButton;
            this.chkEmbedStandardToolbar.Checked = GlobalSettings.UseStandardToolbar;
            this.chkReportToAllProfiles.Checked = GlobalSettings.ReportToMultipleProfiles;
            this.chkShowSupportButton.Checked = GlobalSettings.ShowSupportButton;
            UIControl.StartWindowUpdating(this.Handle);
            this.Invalidate(true);
            this.Refresh();
        }

        #region Event Handlers

        private void chkShowPreviewButtonBox_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ShowPreviewButton = chkShowPreviewButtonBox.Checked;
        }

        private void chkShowCopyButton_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ShowCopyButton = chkShowCopyButton.Checked;
        }

        private void chkEmbedStandardToolbar_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.UseStandardToolbar = chkEmbedStandardToolbar.Checked;
        }

        private void chkShowHam_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ShowHamButton = chkShowHamButton.Checked;
        }

        private void ctlGeneralOptions_Load(object sender, EventArgs e)
        {
            ToolTip ttGeneralOptions = new ToolTip();

            // Set up the delays for the ToolTip.
            ttGeneralOptions.AutoPopDelay = 5000;
            ttGeneralOptions.InitialDelay = 1000;
            ttGeneralOptions.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            ttGeneralOptions.ShowAlways = true;

            // Set the tooltips for the control
            ttGeneralOptions.SetToolTip(chkEmbedStandardToolbar, "Embed the SpamGrabber buttons on the standard Outlook toolbar");
            ttGeneralOptions.SetToolTip(chkReportToAllProfiles, "Tell the application to send reports to all \r\nprofiles in the spam and ham category.");
            ttGeneralOptions.SetToolTip(chkShowCopyButton, "Enable the Copy To Clipboard function.");
            ttGeneralOptions.SetToolTip(chkShowSupportButton, "Enable the Send to support function.");
            ttGeneralOptions.SetToolTip(chkShowSelectButton, "Enable the Copy To Selected profile function.");
            ttGeneralOptions.SetToolTip(chkShowHamButton, "Enable reporting of Ham to a selected profile.");
            ttGeneralOptions.SetToolTip(chkShowPreviewButtonBox, "Show the preview button to enable safe preview of messages.");

        }

        private void chkReportToAllProfiles_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ReportToMultipleProfiles = chkReportToAllProfiles.Checked;
        }

        private void chkShowSelectButton_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ShowSelectButton = chkShowSelectButton.Checked;
        }

        private void chkShowSupportButton_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ShowSupportButton = chkShowSupportButton.Checked;
        }
        #endregion

        

        

        
    }
}
