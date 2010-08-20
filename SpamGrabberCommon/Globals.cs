#region Imports

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.Text.RegularExpressions;

#endregion

namespace SpamGrabberCommon
{
    #region Globals class

    /// <summary>
    /// Constants and code relating to the entire program
    /// </summary>

    public static class SGGlobals
    {
        #region Enums

        public enum BoolRegKey
        {
            True,
            False,
            UnSet
        }

        #endregion

        #region Class Data

        private static string _strBaseRegKey = "Software\\SpamGrabber\\";
        private static string _strEmailRegEx = @"\w+@\w+\.\w+((\.\w+)*)?";

        public enum ReportAction
        {
            ReportSpam,
            ReportHam,
            CopyToClipboard,
            ReportSupport
        }
        #endregion

        #region Properties

        /// <summary>
        /// The registry key where all user settings are stored
        /// </summary>
        public static string BaseRegistryKey
        {
            get
            {
                return _strBaseRegKey;
            }
        }

        /// <summary>
        /// A regular expression for validating email addresses
        /// </summary>
        public static string EmailRegEx
        {
            get
            {
                return _strEmailRegEx;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Returns false if the base settings registry key does not exist
        /// </summary>
        /// <returns></returns>
        public static bool DoesBaseKeyExist()
        {
            // Declare return value
            bool blnKeyExists = false;

            // Get the root HKCU key
            RegistryKey regSettings = Registry.CurrentUser;

            // Try and open the main key
            try
            {
                regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey,true);
                if (regSettings == null)
                {
                    blnKeyExists = false;
                }
                else
                {
                    blnKeyExists = true;
                }
            }
            catch (NullReferenceException)
            {
                blnKeyExists = false;
            }
            catch (Exception)
            {
                blnKeyExists = false;
            }
            return blnKeyExists;
        }

        public static bool DoesKeyExist(string pstrKey)
        {
            // Declare return value
            bool blnKeyExists = false;

            // Get the root HKCU key
            RegistryKey regSettings = Registry.CurrentUser;

            // Try and open the main key
            try
            {
                regSettings = regSettings.OpenSubKey(pstrKey, true);
                if (regSettings == null)
                {
                    blnKeyExists = false;
                }
                else
                {
                    blnKeyExists = true;
                }
            }
            catch (NullReferenceException)
            {
                blnKeyExists = false;
            }
            catch (Exception)
            {
                blnKeyExists = false;
            }
            return blnKeyExists;
        }

        /// <summary>
        /// Creates the base settings key and required subkeys
        /// </summary>
        public static void CreateBaseKey()
        {
            // Get the root HKCU key
            RegistryKey regSettings = Registry.CurrentUser;

            // Create the key
            if (DoesBaseKeyExist() == false)
            {
                // Create the base key
                regSettings = regSettings.CreateSubKey(SGGlobals.BaseRegistryKey);

                // Create the profiles subkey
                regSettings = regSettings.CreateSubKey("Profiles");
            }
        }

        /// <summary>
        /// Creates the settings key and required parentkey
        /// only supports one level
        /// </summary>
        public static void CreateKey(string pstrKey, string pstrParent)
        {
            // Get the root HKCU key
            RegistryKey regSettings = Registry.CurrentUser;

            // Create the key
            if (DoesKeyExist(pstrParent) == false)
            {
                // Create the base key
                regSettings = regSettings.CreateSubKey(pstrParent);

                // Create the profiles subkey
                regSettings = regSettings.CreateSubKey(pstrKey);
            }
            else
            {
                if (DoesKeyExist(pstrKey) == false)
                {
                    regSettings = regSettings.OpenSubKey(pstrParent,true);
                    regSettings = regSettings.CreateSubKey(pstrKey);
                }
            }
        }

        /// <summary>
        /// Loads the string setting from the database, creating it if it
        /// does not already exist
        /// </summary>
        /// <param name="pregSettingsKey">A RegistryKey object pointing to the correct registry key</param>
        /// <param name="pstrSettingName">The setting name to retrieve</param>
        /// <param name="pblnDefault">The default value to use if the key needs to be created</param>
        /// <returns></returns>
        public static bool LoadValue(RegistryKey pregSettingsKey, string pstrSettingName, bool pblnDefault)
        {
            if (pregSettingsKey.GetValue(pstrSettingName) == null)
            {
                // Create the key
                //pregSettingsKey.CreateSubKey(pstrSettingName);
                pregSettingsKey.SetValue(pstrSettingName, pblnDefault, RegistryValueKind.DWord);
                return pblnDefault;
            }
            else
            {
                // Load the key
                if (((Int32)pregSettingsKey.GetValue(pstrSettingName)) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// Loads the string setting from the database, creating it if it
        /// does not already exist
        /// </summary>
        /// <param name="pregSettingsKey">A RegistryKey object pointing to the correct registry key</param>
        /// <param name="pstrSettingName">The setting name to retrieve</param>
        /// <param name="pblnDefault">The default value to use if the key needs to be created</param>
        /// <returns></returns>
        public static Int32 LoadValue(RegistryKey pregSettingsKey, string pstrSettingName, Int32 piDefault)
        {
            if (pregSettingsKey.GetValue(pstrSettingName) == null)
            {
                // Create the key
                //pregSettingsKey.CreateSubKey(pstrSettingName);
                pregSettingsKey.SetValue(pstrSettingName, piDefault, RegistryValueKind.DWord);
                return piDefault;
            }
            else
            {
                // Load the key
                return (Int32)pregSettingsKey.GetValue(pstrSettingName);

            }
        }

        /// <summary>
        /// Loads the boolean setting from the database, creating it if it does not
        /// already exist
        /// </summary>
        /// <param name="pregSettingsKey">A RegistryKey object pointing to the correct registry key</param>
        /// <param name="pstrSettingName">The setting name to retrieve</param>
        /// <param name="pstrDefault">The default value to use if the key needs to be created</param>
        /// <returns></returns>
        public static string LoadValue(RegistryKey pregSettingsKey, string pstrSettingName, string pstrDefault)
        {
            if (pregSettingsKey.GetValue(pstrSettingName) == null)
            {
                // Create the key
                //pregSettingsKey.CreateSubKey(pstrSettingName);
                pregSettingsKey.SetValue(pstrSettingName, pstrDefault, RegistryValueKind.String);
                return pstrDefault;
            }
            else
            {
                return (string)pregSettingsKey.GetValue(pstrSettingName);
            }
        }

        public static Int32 LoadValue(string pstrKey, string pstrSettingName, Int32 pintDefault)
        {
            RegistryKey regSettings = Registry.CurrentUser;

            regSettings = regSettings.OpenSubKey(pstrKey, true);

            if (regSettings == null)
            {
                CreateBaseKey();
                regSettings = Registry.CurrentUser;
                regSettings = regSettings.OpenSubKey(pstrKey, true);
            }

            if (regSettings.GetValue(pstrSettingName) == null)
            {
                regSettings.SetValue(pstrSettingName, pintDefault, RegistryValueKind.DWord);
                return pintDefault;
            }
            else
            {
                return (Int32)regSettings.GetValue(pstrSettingName, pintDefault);

            }

        }

        public static string LoadValue(string pstrKey, string pstrSettingName, string pstrDefault)
        {
            RegistryKey regSettings = Registry.CurrentUser;


            regSettings = regSettings.OpenSubKey(pstrKey,true);

            if (regSettings == null)
            {
                CreateBaseKey();
                regSettings = Registry.CurrentUser;
                regSettings = regSettings.OpenSubKey(pstrKey, true);
            }

            if (regSettings.GetValue(pstrSettingName) == null)
            {
                regSettings.SetValue(pstrSettingName, pstrDefault, RegistryValueKind.String);
                return pstrDefault;
            }
            else
            {
                return (string)regSettings.GetValue(pstrSettingName, pstrDefault);
            }

        }

        public static bool LoadValue(string pstrKey, string pstrSettingName, bool pblnDefault)
        {

            Int32 intReturn;

            RegistryKey regSettings = Registry.CurrentUser;

            regSettings = regSettings.OpenSubKey(pstrKey,true);

            if (regSettings == null)
            {
                CreateBaseKey();
                regSettings = Registry.CurrentUser;
                regSettings = regSettings.OpenSubKey(pstrKey, true);
            }


            if (regSettings.GetValue(pstrSettingName) == null)
            {
                if (pblnDefault == true)
                    regSettings.SetValue(pstrSettingName, true, RegistryValueKind.DWord);
                else
                    regSettings.SetValue(pstrSettingName, false, RegistryValueKind.DWord);
                return pblnDefault;
            }
            else
            {
                intReturn = (Int32)regSettings.GetValue(pstrSettingName, pblnDefault);
                if (intReturn == 0)
                    return false;
                else
                    return true;
            }
        }

        /// <summary>
        /// Verifies an email address format
        /// </summary>
        /// <param name="pstrEmail"></param>
        /// <returns></returns>
        public static bool IsEmailValid(string pstrEmail)
        {
            return Regex.Match(pstrEmail, EmailRegEx).Success;
        }

        ///// <summary>
        ///// Gets the requested boolean setting from the registry, creating it with the
        ///// default of true if it does not exist
        ///// </summary>
        ///// <param name="pstrSettingName">The setting name as held in the registry</param>
        ///// <returns></returns>
        //public static bool GetSettingValue(string pstrSettingName)
        //{
        //    // Get the root HKCU key
        //    RegistryKey regSettings = Registry.CurrentUser;

        //    // Get the root SpamGrabber key
        //    regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey, true);

        //    // Make sure the key exists
        //    if (regSettings == null)
        //    {
        //        // Create the default keys
        //        SGGlobals.CreateBaseKey();
        //        // and try again
        //        regSettings = Registry.CurrentUser;
        //        regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey, true);
        //    }

        //    // Return the selected value
        //    return SGGlobals.LoadValue(regSettings, pstrSettingName, true);
        //}

        ///// <summary>
        ///// Gets the requested string setting from the registry, creating it with the
        ///// supplied default if it does not exist
        ///// </summary>
        ///// <param name="pstrSettingName"></param>
        ///// <param name="pstrDefault"></param>
        ///// <returns></returns>
        //public static string GetSettingValue(string pstrSettingName, string pstrDefault)
        //{
        //    // Get the root HKCU key
        //    RegistryKey regSettings = Registry.CurrentUser;

        //    // Get the root SpamGrabber key
        //    regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey, true);

        //    // Make sure the key exists
        //    if (regSettings == null)
        //    {
        //        // Create the default keys
        //        SGGlobals.CreateBaseKey();
        //        // and try again
        //        regSettings = Registry.CurrentUser;
        //        regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey, true);
        //    }

        //    // Return the selected value
        //    return SGGlobals.LoadValue(regSettings, pstrSettingName, pstrDefault);

        //}

        /// <summary>
        /// Saves the specified boolean setting to the database
        /// </summary>
        /// <param name="pstrSettingName"></param>
        /// <param name="pblnSettingValue"></param>
        public static void SaveSetting(string pstrKey, string pstrSettingName, bool pblnSettingValue)
        {
            // Save the setting (should already exist, provided they have
            // called this function using the property)

            // Open the base settings key
            RegistryKey regSettings = Registry.CurrentUser.OpenSubKey(pstrKey, true);

            // Save the value
            regSettings.SetValue(pstrSettingName, pblnSettingValue, RegistryValueKind.DWord);
        }

        /// <summary>
        /// Saves the specified Int32 setting to the database
        /// </summary>
        /// <param name="pstrSettingName"></param>
        /// <param name="pblnSettingValue"></param>
        public static void SaveSetting(string pstrKey, string pstrSettingName, Int32 pblnSettingValue)
        {
            // Save the setting (should already exist, provided they have
            // called this function using the property)

            // Open the base settings key
            RegistryKey regSettings = Registry.CurrentUser.OpenSubKey(pstrKey, true);

            // Save the value
            regSettings.SetValue(pstrSettingName, pblnSettingValue, RegistryValueKind.DWord);
        }

        /// <summary>
        /// Saves the specified string setting to the database
        /// </summary>
        /// <param name="pstrSettingName"></param>
        /// <param name="pstrSettingValue"></param>
        public static void SaveSetting(string pstrKey, string pstrSettingName, string pstrSettingValue)
        {
            // Save the setting (should already exist, provided they have
            // called this function using the property)

            // Open the base settings key
            RegistryKey regSettings = Registry.CurrentUser.OpenSubKey(pstrKey, true);

            // Save the value
            regSettings.SetValue(pstrSettingName, pstrSettingValue, RegistryValueKind.String);
        }

        public static string GetUserConfigurationKey(string pstrKey)
        {
            // New Microsoft.Win32 Registry Key 
            RegistryKey regkey = null;
            // Path to top key
            regkey = Registry.CurrentUser.OpenSubKey(pstrKey, true);

            if (regkey == null)
                return "";
            else
                return regkey.ToString();
        }

        public static string GetBaseConfigurationKey()
        {
            return BaseRegistryKey;
        }

        public static string CreateMachineConfigurationKey(string pstrKey)
        {
            // New Microsoft.Win32 Registry Key 
            RegistryKey regkey = null;
            regkey = Registry.CurrentUser.OpenSubKey(pstrKey, true);

            if (regkey == null)
                return "";
            else
                return regkey.ToString();
        }

        /// <summary>
        /// the original VB code for repairing headers. This is used if selected by user.
        /// </summary>
        /// <param name="header">The header to fix</param>
        /// <param name="msgIsHTML">Wheter or not it is a HTML mail</param>
        /// <returns>The fixed header</returns>
        public static string RepairHeaders(string pstrHeader, bool pblnMsgIsHTML)
        {
            List<string> lstHeader = new List<string>();
            List<string> tempLines = new List<string>();
            string removeString, removeString2;
            bool headerFlag;


            string[] del = new string[] { "\r\n" };
            lstHeader.AddRange(pstrHeader.Split(new string[] { "\r\n\r\n" }, StringSplitOptions.None));

            string msgHeaders = "";
            string tempHeader = "";
            //' OK, first we need to break off any boundary headers which are
            //' caused by attachments etc. These are easy to get rid of, as
            //' they always appear two CrLfs after the main headers, so
            //' we just do a split on a double CrLf and take the first part
            msgHeaders = lstHeader[0];

            //' Loop through each line in the headers and remove any
            //' Content-type lines or Microsoft headers
            removeString = "Content-Type";
            removeString2 = "Microsoft Mail Internet Headers Version 2.0";
            headerFlag = false;
            tempLines.AddRange(msgHeaders.Split(del, StringSplitOptions.None));

            foreach (string tempLine in tempLines)
            {
                if (tempLine != string.Empty)
                {
                    if (headerFlag)
                    {
                        //' Flag is on, so we need to see if this is a normal header or
                        //' a continuation header
                        if (tempLine.Substring(0, 1) == " " || tempLine.Substring(0, 1) == "\t")
                        { }   // ' It is a continuation, so do nothing
                        else
                        {
                            //' New header, check to make sure it isn't the
                            //' MS one, and add it
                            if (tempLine.Contains(removeString2))
                            {    //' Just add the header
                                tempHeader += tempLine + "\r\n";
                            }
                            //' Reset the flag
                            headerFlag = false;
                        }
                    }
                    else
                    {  //' Flag is off, so check to see if this is a content type header
                        if (tempLine.Contains(removeString))
                        {    //' This is a content type header, so set the flag and ignore the header
                            headerFlag = true;
                        }
                        else
                        {
                            //' Not a content type header, so check for the MS header and add
                            if (!tempLine.Contains(removeString2))
                            {
                                //' Just add the header
                                tempHeader += tempLine + "\r\n";
                            }
                        }
                    }
                }
            }

            //' Add the correct content type header at the end
            tempHeader = tempHeader + "Content-Type: text/";
            if (pblnMsgIsHTML)
            {
                tempHeader = tempHeader + "html;";
            }
            else
            {
                tempHeader = tempHeader + "plain;";
            }

            //' Set the reporting headers to the cleaned headers
            return tempHeader + "\r\n\r\n";
        }

        #endregion
    }

    #endregion
}
