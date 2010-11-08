using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Explorer _objExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            _cbSpamGrabber = _objExplorer.CommandBars.Add("SpamGrabber", Office.MsoBarPosition.msoBarTop, false, true);

            _cbbDefaultSpam = CreateCommandBarButton(_cbSpamGrabber,
                "Report Spam", "Report to Default Spam profile", "Report to Default Spam profile",
                1, Office.MsoButtonStyle.msoButtonIconAndCaption,
                true, true, 1);

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

        public static Office.CommandBarButton CreateCommandBarButton(
            Office.CommandBar commandBar, string captionText, string tagText, 
            string tipText, int faceID, Office.MsoButtonStyle buttonStyle, 
            bool beginGroup, bool isVisible, object objBefore)
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
                    aButton.Picture = (IPictureDisp)GetIPictureDispFromPicture(Properties.Resources.spamgrab_red);
                }
                aButton.Style = buttonStyle;
                aButton.TooltipText = tipText;
                aButton.BeginGroup = beginGroup;
                //aButton.OnAction = ProgID;

            }

            aButton.Visible = isVisible;

            return aButton;
        }
    }
}
