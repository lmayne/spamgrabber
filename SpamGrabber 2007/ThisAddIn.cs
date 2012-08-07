using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using stdole;
using System.Windows.Forms;
using SpamGrabberCommon;
using SpamGrabberControl;

namespace SpamGrabber_2007
{
    public partial class ThisAddIn
    {
        private Office.CommandBar _cbSpamGrabber;
        private Office.CommandBarButton _cbbDefaultHam;
        private Office.CommandBarButton _cbbDefaultSpam;
        private Office.CommandBarButton _cbbCopyToClipboard;
        private Office.CommandBarButton _cbbPreview;
        private Office.CommandBarButton _cbbOptions;
        private Office.CommandBarComboBox _cbcbProfile;
        private Office.CommandBarButton _cbbReportSelected;
        private Office._CommandBarButtonEvents_ClickEventHandler _ReportSpam;
        private Office._CommandBarButtonEvents_ClickEventHandler _ReportHam;
        private Office._CommandBarButtonEvents_ClickEventHandler _ReportSelected;
        private Office._CommandBarButtonEvents_ClickEventHandler _SafeView;
        private Office._CommandBarButtonEvents_ClickEventHandler _CopyToClipboard;
        private Office._CommandBarButtonEvents_ClickEventHandler _Settings;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _ReportSpam = ReportSpam;
            _ReportHam = ReportHam;
            _ReportSelected = ReportSelected;
            _SafeView = SafeView;
            _CopyToClipboard = CopyToClipboard;
            _Settings = Settings;

            Explorer _objExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            //_cbSpamGrabber = _objExplorer.CommandBars.Add("SpamGrabber", Office.MsoBarPosition.msoBarTop, false, true);
            _cbSpamGrabber = _objExplorer.CommandBars["Standard"];

            _cbbDefaultSpam = CreateCommandBarButton(_cbSpamGrabber,
                "Report Spam", "Report to Default Spam profile", "Report to Default Spam profile",
                Office.MsoButtonStyle.msoButtonIcon, Properties.Resources.spamgrab_red,
                true, true, _cbSpamGrabber.Controls.Count, _ReportSpam);

            _cbbDefaultHam = CreateCommandBarButton(_cbSpamGrabber,
                "Report Ham", "Report to Default Ham profile", "Report to Default Ham profile",
                Office.MsoButtonStyle.msoButtonIcon, Properties.Resources.spamgrab_green,
                false, GlobalSettings.ShowHamButton, _cbSpamGrabber.Controls.Count, _ReportHam);

            _cbcbProfile = AddComboBox(_cbSpamGrabber);
            _cbcbProfile.Visible = GlobalSettings.ShowSelectButton;

            _cbbReportSelected = CreateCommandBarButton(_cbSpamGrabber,
                "Report", "Report to Selected Profile", "Report to Selected Profile",
                Office.MsoButtonStyle.msoButtonCaption, null,
                false, GlobalSettings.ShowSelectButton, _cbSpamGrabber.Controls.Count, _ReportSelected);

            _cbbPreview = CreateCommandBarButton(_cbSpamGrabber,
                "Safe View", "Safe Preview", "Safe Preview",
                Office.MsoButtonStyle.msoButtonIcon, Properties.Resources.search4doc,
                true, GlobalSettings.ShowPreviewButton, _cbSpamGrabber.Controls.Count, _SafeView);

            _cbbCopyToClipboard = CreateCommandBarButton(_cbSpamGrabber,
                "Copy Source", "Copy Source to Clipboard", "Copy Source to Clipboard",
                Office.MsoButtonStyle.msoButtonIcon, Properties.Resources.spamgrab_copy,
                false, GlobalSettings.ShowCopyButton, _cbSpamGrabber.Controls.Count, _CopyToClipboard);

            _cbbOptions = CreateCommandBarButton(_cbSpamGrabber,
                "Options", "SpamGrabber Options", "SpamGrabber Options",
                Office.MsoButtonStyle.msoButtonIcon, Properties.Resources.spamgrab_settings,
                false, GlobalSettings.ShowSettingsButton, _cbSpamGrabber.Controls.Count, _Settings);

            Reporting.Application = Globals.ThisAddIn.Application;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

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

        private Office.CommandBarButton CreateCommandBarButton(
        Office.CommandBar commandBar, string captionText, string tagText,
        string tipText, Office.MsoButtonStyle buttonStyle, System.Drawing.Bitmap picture,
        bool beginGroup, bool isVisible, object objBefore, Office._CommandBarButtonEvents_ClickEventHandler handler)
        {
            // Determine if button exists
            Office.CommandBarButton aButton = (Office.CommandBarButton)
                commandBar.FindControl(buttonStyle, null, tagText, null, null);

            if (aButton == null)
            {
                // Add new button
                aButton = (Office.CommandBarButton)
                    commandBar.Controls.Add(Office.MsoControlType.msoControlButton, 1, tagText, objBefore, true);

                aButton.Caption = captionText;
                aButton.Tag = tagText;
                if (buttonStyle != Office.MsoButtonStyle.msoButtonCaption)
                {
                    aButton.Picture = (IPictureDisp)AxHost2.GetIPictureDispFromPicture(picture);
                }
                aButton.Style = buttonStyle;
                aButton.TooltipText = tipText;
                aButton.BeginGroup = beginGroup;
                aButton.Click += handler;
            }

            aButton.Visible = isVisible;

            return aButton;
        }

        private Office.CommandBarComboBox AddComboBox(Office.CommandBar commandBar)
        {
            if (_cbcbProfile == null)
            {
                _cbcbProfile = (Office.CommandBarComboBox)
                    commandBar.Controls.Add(Office.MsoControlType.msoControlComboBox, 1, "Select spam profile", commandBar.Controls.Count, true);

                _cbcbProfile.Style = Office.MsoComboStyle.msoComboLabel;
                _cbcbProfile.Caption = "Select profile:";
                _cbcbProfile.TooltipText = "Select profile to report to";
                _cbcbProfile.BeginGroup = true;
                this.LoadDropDown();
            }
            return _cbcbProfile;
        }

        private void ReportSpam(Office.CommandBarButton btn, ref bool cancel)
        {
            if (string.IsNullOrEmpty(SpamGrabberCommon.GlobalSettings.DefaultSpamProfileId))
            {
                MessageBox.Show("You have not yet set a default spam profile. Please open the SpamGrabber settings dialog and set a default spam profile");
                return;
            }
            Reporting.SendReports(SpamGrabberCommon.GlobalSettings.DefaultSpamProfileId);
        }

        private void ReportHam(Office.CommandBarButton btn, ref bool cancel)
        {
            if (string.IsNullOrEmpty(SpamGrabberCommon.GlobalSettings.DefaultHamProfileId))
            {
                MessageBox.Show("You have not yet set a default ham profile. Please open the SpamGrabber settings dialog and set a default ham profile");
                return;
            }
            Reporting.SendReports(SpamGrabberCommon.GlobalSettings.DefaultHamProfileId);
        }

        private void ReportSelected(Office.CommandBarButton btn, ref bool cancel)
        {
            try
            {
                Profile objProfile = SpamGrabberCommon.UserProfiles.GetProfileByName(this._cbcbProfile.Text);
                if (objProfile != null)
                {
                    Reporting.SendReports(objProfile.Id);
                }
            }
            catch (ProfileNotFoundException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void SafeView(Office.CommandBarButton btn, ref bool cancel)
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
        
        private void CopyToClipboard(Office.CommandBarButton btn, ref bool cancel)
        {
            Explorer exp = Globals.ThisAddIn.Application.ActiveExplorer();
            if (exp.Selection.Count > 0)
            {
                Clipboard.SetText(Reporting.GetMessageSource((MailItem)exp.Selection[1], false));
            }
        }

        private void Settings(Office.CommandBarButton btn, ref bool cancel)
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

        private void LoadDropDown()
        {
            if (_cbcbProfile != null)
            {
                _cbcbProfile.Clear();
                foreach (SpamGrabberCommon.Profile profile in SpamGrabberCommon.UserProfiles.ProfileList)
                {
                    _cbcbProfile.AddItem(profile.Name);
                }
            }
        }

        private void ShowHideButtons()
        {
            this._cbbDefaultHam.Visible = SpamGrabberCommon.GlobalSettings.ShowHamButton;
            this._cbbCopyToClipboard.Visible = SpamGrabberCommon.GlobalSettings.ShowCopyButton;
            this._cbbPreview.Visible = SpamGrabberCommon.GlobalSettings.ShowPreviewButton;
            this._cbbOptions.Visible = SpamGrabberCommon.GlobalSettings.ShowSettingsButton;
            this._cbcbProfile.Visible = SpamGrabberCommon.GlobalSettings.ShowSelectButton;
            this._cbbReportSelected.Visible = SpamGrabberCommon.GlobalSettings.ShowSelectButton;
        }
    }
}
