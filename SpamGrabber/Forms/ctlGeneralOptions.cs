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
            ttGeneralOptions.SetToolTip(chkShowCopyButton, "Enable the Copy To Clipboard function.");
            ttGeneralOptions.SetToolTip(chkShowSelectButton, "Enable the Copy To Selected profile function.");
            ttGeneralOptions.SetToolTip(chkShowHamButton, "Enable reporting of Ham to a selected profile.");
            ttGeneralOptions.SetToolTip(chkShowPreviewButtonBox, "Show the preview button to enable safe preview of messages.");

        }

        private void chkShowSelectButton_CheckedChanged(object sender, EventArgs e)
        {
            GlobalSettings.ShowSelectButton = chkShowSelectButton.Checked;
        }

        #endregion

        

        

        
    }
}
