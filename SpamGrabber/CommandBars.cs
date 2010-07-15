#region imports

using System;
using Office = Microsoft.Office.Core;
using stdole;
using System.Windows.Forms;
using System.Drawing;
using System.Resources;
using SpamGrabberCommon;

#endregion

namespace SpamGrabberCommon
{
    /// <summary>
    /// Helper class to create command bar buttons
    /// </summary>
    class CommandBars
    {
        #region Methods

        /// <summary>
        /// Adds the selected button to the specified command bar
        /// </summary>
        /// <param name="commandBar">The command bar to add the button to</param>
        /// <param name="captionText">The text displayed in the button</param>
        /// <param name="tagText">The tag name of the button</param>
        /// <param name="tipText">The tooltip text of the button</param>
        /// <param name="faceID">ID number of the icon to use</param>
        /// <param name="buttonStyle">Button stle to use (e.g. icon only, icon and text)</param>
        /// <param name="beginGroup">Whether or not to add a new spacer before the button</param>
        /// <param name="isVisible">Whether the button should be displayed or not</param>
        /// <param name="objBefore">The button in the command bar to place the item before</param>
        /// <returns></returns>
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
                    aButton.FaceId = faceID;
                }
                aButton.Style = buttonStyle;
                aButton.TooltipText = tipText;
                aButton.BeginGroup = beginGroup;
                //aButton.OnAction = ProgID;

            }

            aButton.Visible = isVisible;

            return aButton;
        }

        public static Office.CommandBarButton CreateCommandBarButton(
            Office.CommandBar commandBar, string captionText, string tagText,
            string tipText, string imageName,string imageMask, Office.MsoButtonStyle buttonStyle,
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
                    ResourceManager rm = SpamGrabber2003.Properties.Resources.ResourceManager;

                    aButton.Picture = MyAxHost.GettIPictureDispFromPicture((Image)rm.GetObject(imageName));
                    // get the mask
                    aButton.Mask = MyAxHost.GettIPictureDispFromPicture((Image)rm.GetObject(imageMask));
                }
                aButton.Style = buttonStyle;
                aButton.TooltipText = tipText;
                aButton.BeginGroup = beginGroup;
                //aButton.OnAction = ProgID;


                
            }

            aButton.Visible = isVisible;

            return aButton;
        }

        public static void LoadCommandBarSettings(Microsoft.Office.Core.CommandBar pobjcommandBar)
        {
            Microsoft.Office.Core.MsoBarPosition position =
                (Microsoft.Office.Core.MsoBarPosition)GlobalSettings.CommandBarPosition;

            pobjcommandBar.Position = position;

            int rowIndex =Convert.ToInt32(GlobalSettings.CommandBarRowIndex);

            pobjcommandBar.RowIndex = rowIndex;

            int top =Convert.ToInt32(GlobalSettings.CommandBarTop);

            pobjcommandBar.Top = top;//!= 0 ? top : Screen.PrimaryScreen.Bounds.Height / 2;

            int left =Convert.ToInt32(GlobalSettings.CommandBarLeft);

            pobjcommandBar.Left = left;// != 0 ? left : Screen.PrimaryScreen.Bounds.Width / 2;

            bool visible = Convert.ToBoolean(GlobalSettings.CommandBarVisible);

            pobjcommandBar.Visible = visible;
        }

        public static void SaveCommandBarSettings(Microsoft.Office.Core.CommandBar pobjcommandBar)
        {

            GlobalSettings.Save();
        }


        #endregion
    }

    class MyAxHost : AxHost
    {
        public MyAxHost()
            : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33")
        {
        }

        public static IPictureDisp GettIPictureDispFromPicture(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }

}
