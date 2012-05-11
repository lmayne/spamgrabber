using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using SpamGrabberControl;
using SpamGrabberCommon;

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
            Reporting.SendReports(SpamGrabberCommon.GlobalSettings.DefaultSpamProfileId);
        }

        private void btnReportDefaultHam_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(SpamGrabberCommon.GlobalSettings.DefaultHamProfileId))
            {
                MessageBox.Show("You have not yet set a default ham profile. Please open the SpamGrabber settings dialog and set a default ham profile");
                return;
            }
            Reporting.SendReports(SpamGrabberCommon.GlobalSettings.DefaultHamProfileId);
        }

        private void btnCopyToClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            Explorer exp = Globals.ThisAddIn.Application.ActiveExplorer();
            if (exp.Selection.Count > 0)
            {
                Clipboard.SetText(Reporting.GetMessageSource((MailItem)exp.Selection[1], false));
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
                    // Refresh the drop down items
                    this.LoadDropDown();
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
                Reporting.SendReports(this.ddlReportTo.SelectedItem.Tag.ToString());
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

        private void SpamGrabber_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Set the application ref in the common class
            Reporting.Application = Globals.ThisAddIn.Application;
            // Load the drop down items
            this.LoadDropDown();
            // Show / hide buttons based on settings
            this.ShowHideButtons();
        }

        private void LoadDropDown()
        {
            this.ddlReportTo.Items.Clear();
            foreach (SpamGrabberCommon.Profile profile in SpamGrabberCommon.UserProfiles.ProfileList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = profile.Name;
                item.Tag = profile.Id;
                this.ddlReportTo.Items.Add(item);
            }
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
