using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using stdole;

namespace SpamGrabber_2007
{
    public partial class ThisAddIn
    {
        private Office.CommandBar _cbSpamGrabber;
        private Office.CommandBarButton _cbbDefaultHam;
        private Office.CommandBarButton _cbbDefaultSpam;
        private Office.CommandBarButton _cbbReportToProfile;
        private Office.CommandBarButton _cbbCopyToClipboard;
        private Office.CommandBarButton _cbbSendToSupport;
        private Office.CommandBarButton _cbbPreview;
        private Office.CommandBarButton _cbbOptions;
        private Office._CommandBarButtonEvents_ClickEventHandler _ReportSpam;
        private Office._CommandBarButtonEvents_ClickEventHandler _ReportHam;
        private Office._CommandBarButtonEvents_ClickEventHandler _SafeView;
        private Office._CommandBarButtonEvents_ClickEventHandler _CopyToClipboard;
        private Office._CommandBarButtonEvents_ClickEventHandler _Settings;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _ReportSpam = ReportSpam;
            _ReportHam = ReportHam;
            _SafeView = SafeView;
            _CopyToClipboard = CopyToClipboard;
            _Settings = Settings;

            Explorer _objExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            _cbSpamGrabber = _objExplorer.CommandBars.Add("SpamGrabber", Office.MsoBarPosition.msoBarTop, false, true);

            _cbbDefaultSpam = CreateCommandBarButton(_cbSpamGrabber,
                "Report Spam", "Report to Default Spam profile", "Report to Default Spam profile",
                Office.MsoButtonStyle.msoButtonIconAndCaption, Properties.Resources.spamgrab_red,
                true, true, 1, _ReportSpam);

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

                aButton.Style = buttonStyle;

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

        private void ReportSpam(Office.CommandBarButton btn, ref bool cancel)
        {
            System.Windows.Forms.MessageBox.Show("Report spam");
        }
        private void ReportHam(Office.CommandBarButton btn, ref bool cancel)
        {

        }
        private void SafeView(Office.CommandBarButton btn, ref bool cancel)
        {

        }
        private void CopyToClipboard(Office.CommandBarButton btn, ref bool cancel)
        {

        }
        private void Settings(Office.CommandBarButton btn, ref bool cancel)
        {

        }
    }
}
