#region Imports

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;

#endregion

namespace SpamGrabberCommon
{
    /// <summary>
    /// Helper class for settings that exist
    /// outside of individual profiles
    /// </summary>
    public class GlobalSettings
    {
        #region Class Data

        private static bool? _blnShowPreviewButton = null;
        private static bool? _blnShowCopyButton = null;
        private static bool? _blnShowHamButton = null;
        private static bool? _blnShowSelectButton = null;
        private static bool? _blnSuppressConfirm = null;
        private static bool? _blnShowSettings = null; // Not available through the GUI
        private static string _strDefaultSpamProfileId = null;
        private static string _strDefaultHamProfileId = null;

        //' Name of commandbar
        public const string COMMANDBAR_NAME = "SpamGrabber";

        // commandbar position information        
        private static Int32 _intCommandBarLeft;
        private static Int32 _intCommandBarPosition;
        private static Int32 _intCommandBarRowIndex;
        private static Int32 _intCommandBarTop;
        private static bool _intCommandBarVisible;




        //' Name of program
        public const string PROGRAM_NAME = "Outlook Spam Report Utility";

        //' Program Version
        public static string PROG_VERSION = "4.0.8";

        public enum DefaultType
        {
            Spam = 0,
            Ham = 1
        }

        #endregion

        #region Properties


        public static int CommandBarLeft
        {
            get
            {

                _intCommandBarLeft = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "CommandBarLeft", 0);

                return GlobalSettings._intCommandBarLeft;
            }

            set { GlobalSettings._intCommandBarLeft = value; }
        }

        public static int CommandBarPosition
        {
            get
            {

                _intCommandBarPosition = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "CommandBarPosition", 1);


                return GlobalSettings._intCommandBarPosition;
            }

            set { GlobalSettings._intCommandBarPosition = value; }
        }

        public static int CommandBarRowIndex
        {
            get
            {

                _intCommandBarRowIndex = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "CommandBarRowIndex", 4);

                return GlobalSettings._intCommandBarRowIndex;
            }

            set { GlobalSettings._intCommandBarRowIndex = value; }
        }

        public static int CommandBarTop
        {
            get
            {

                _intCommandBarTop = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "CommandBarTop", 0);

                return GlobalSettings._intCommandBarTop;
            }

            set { GlobalSettings._intCommandBarTop = value; }
        }

        public static bool CommandBarVisible
        {
            get
            {

                _intCommandBarVisible = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "CommandBarWidth", true);

                return GlobalSettings._intCommandBarVisible;
            }

            set { GlobalSettings._intCommandBarVisible = value; }
        }

        /// <summary>
        /// Whether or not the Preview Message button should be displayed
        /// </summary>
        public static bool ShowPreviewButton
        {
            get
            {
                // If the value has not been populated...
                if (_blnShowPreviewButton == null)
                {
                    // ...get it from the registry
                    _blnShowPreviewButton = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "ShowPreviewButton", false);
                }
                return (bool)_blnShowPreviewButton;
            }
            set
            {
                _blnShowPreviewButton = value;
            }
        }

        /// <summary>
        /// Whether or not the Copy To Clipboard button should be displayed
        /// </summary>
        public static bool ShowCopyButton
        {
            get
            {
                // If the value has not been populated...
                if (_blnShowCopyButton == null)
                {
                    // ...get it from the registry
                    _blnShowCopyButton = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "ShowCopyButton", true);
                }
                return (bool)_blnShowCopyButton;
            }
            set
            {
                _blnShowCopyButton = value;
            }
        }

        /// <summary>
        /// Whether or not the default ham button should be displayed
        /// </summary>
        public static bool ShowHamButton
        {
            get
            {
                // If the value has not been populated...
                if (_blnShowHamButton == null)
                {
                    // ...get it from the registry
                    _blnShowHamButton = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "ShowHamButton", false);
                }
                return (bool)_blnShowHamButton;
            }
            set
            {
                _blnShowHamButton = value;
            }
        }

        /// <summary>
        /// Whether or not the BSelect button should be displayed
        /// </summary>
        public static bool ShowSelectButton
        {
            get
            {
                // If the value has not been populated...
                if (_blnShowSelectButton == null)
                {
                    // ...get it from the registry
                    _blnShowSelectButton = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "ShowSelectButton", false);
                }
                return (bool)_blnShowSelectButton;
            }
            set
            {
                _blnShowSelectButton = value;
            }
        }

        /// <summary>
        /// Suppresses the 'are you sure' message when copying to the clipboard
        /// </summary>
        public static bool SuppressConfirm
        {
            get
            {
                // If the value has not been populated...
                if (_blnSuppressConfirm == null)
                {
                    // ...get it from the registry
                    _blnSuppressConfirm = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "SuppressConfirm", true);
                }
                return (bool)_blnSuppressConfirm;
            }
            set
            {
                _blnSuppressConfirm = value;
            }
        }

        /// <summary>
        /// Whether or not to display the edit settings button
        /// </summary>
        public static bool ShowSettingsButton
        {
            get
            {
                // If the value has not been populated...
                if (_blnShowSettings == null)
                {
                    // ...get it from the registry
                    _blnShowSettings = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "ShowSettingsButton", true);
                }
                return (bool)_blnShowSettings;
            }
            set
            {
                _blnShowSettings = value;
            }
        }

        public static string DefaultSpamProfileId
        {
            get
            {
                // If the setting has not been populated...
                if (_strDefaultSpamProfileId == null)
                {
                    // ...get it from the registry
                    _strDefaultSpamProfileId = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "DefaultSpamProfileId", "");
                }
                return _strDefaultSpamProfileId;
            }
            set
            {
                _strDefaultSpamProfileId = value;
            }
        }

        public static string DefaultHamProfileId
        {
            get
            {
                // If the setting has not been populated...
                if (_strDefaultHamProfileId == null)
                {
                    // ...get it from the registry
                    _strDefaultHamProfileId = SGGlobals.LoadValue(SGGlobals.BaseRegistryKey, "DefaultHamProfileId", "");
                }
                return _strDefaultHamProfileId;
            }
            set
            {
                _strDefaultHamProfileId = value;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Saves the stored data
        /// </summary>
        public static void Save()
        {
            // Save all the settings one by one
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "ShowPreviewButton", ShowPreviewButton);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "ShowCopyButton", ShowCopyButton);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "ShowHamButton", ShowHamButton);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "ShowSelectButton", ShowSelectButton);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "SuppressConfirm", SuppressConfirm);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "DefaultSpamProfileId", DefaultSpamProfileId);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "DefaultHamProfileId", DefaultHamProfileId);

            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "CommandBarLeft", _intCommandBarLeft);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "CommandBarPosition", _intCommandBarPosition);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "CommandBarRowIndex", _intCommandBarRowIndex);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "CommandBarTop", _intCommandBarTop);
            SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "CommandBarVisible", _intCommandBarVisible);

        }

        /// <summary>
        /// Re-populates all the settings from the database, discarding any changes made
        /// </summary>
        public static void LoadSettings()
        {
            string strKey = SGGlobals.BaseRegistryKey;

            // Make sure the key exists
            if (!SGGlobals.DoesKeyExist(strKey))
            {
                throw new Exception("The confguration key could not be found");
            }

            _blnShowPreviewButton = SGGlobals.LoadValue(strKey, "ShowPreviewButton", true);
            _blnShowCopyButton = SGGlobals.LoadValue(strKey, "ShowCopyButton", true);
            _blnShowHamButton = SGGlobals.LoadValue(strKey, "ShowHamButton", true);
            _blnShowSelectButton = SGGlobals.LoadValue(strKey, "ShowSelectButton", false);
            _blnSuppressConfirm = SGGlobals.LoadValue(strKey, "SuppressConfirm", true);
        }

        /// <summary>
        /// Resets the default Ham / Spam profile
        /// </summary>
        /// <param name="dftProfile"></param>
        public static void ResetDefaultProfile(DefaultType pdftProfile)
        {
            switch (pdftProfile)
            {
                case DefaultType.Ham:
                    SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "DefaultHamProfileId", string.Empty);
                    break;
                case DefaultType.Spam:
                    SGGlobals.SaveSetting(SGGlobals.BaseRegistryKey, "DefaultSpamProfileId", string.Empty);
                    break;
            }
        }


        #endregion
    }
}
