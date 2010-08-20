#region Imports

using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using Microsoft.Win32;

#endregion

namespace SpamGrabberCommon
{
    #region Profile Class

    /// <summary>
    /// Class symbolising a user profile for SpamGrabber
    /// </summary>
    public class Profile
    {

        #region Class Data

        private string _strProfileName;
        private string _strProfileId;
        private bool? _blnCleanHeaders;
        private List<string> _arrToAddresses;
        private List<string> _arrBccAddresses;
        private bool? _blnDeleteAfterReport;
        private bool? _blnMarkAsReadAfterReport;
        private bool? _blnMoveToFolderAfterReport;
        private string _strMoveFolderName;
        private string _strMoveFolderStoreId;
        private string _strReportSubject;
        private string _strReportEndText;
        private string _strMessageBody;
        private bool? _blnAskVerify;
        private bool? _blnKeepCopy;
        private bool? _blnSendMultiple;
        private bool? _blnUseRFC;
        private bool? _blnSendReceiveAfterReport;

        #endregion

        #region Constructors

        /// <summary>
        /// Contructor for creating a new profile
        /// </summary>
        public Profile()
        {
            // This is a new profile, so assign a GUID
            this._strProfileId = System.Guid.NewGuid().ToString();

            // Load in all the default options and save
            Save();
        }

        /// <summary>
        /// Constructor to load an existing profile from the registry
        /// </summary>
        /// <param name="pstrProfileId"></param>
        public Profile(string pstrProfileId)
        {
            // Set the Profile ID
            this._strProfileId = pstrProfileId;
            // Load the data from the registry
            LoadProfile();
        }

        #endregion

        # region Properties

        /// <summary>
        /// The Outlook folder ID to move messages to, if MoveToFolderAfterReport is true
        /// </summary>
        public string MoveFolderName
        {
            get
            {
                if (this._strMoveFolderName == null)
                {
                    // Default
                    this._strMoveFolderName = "";
                }
                return this._strMoveFolderName;
            }
            set
            {
                this._strMoveFolderName = value;
            }
        }

        /// <summary>
        /// The Outlook folder Store ID to move messages to, if MoveToFolderAfterReport is true
        /// </summary>
        public string MoveFolderStoreId
        {
            get
            {
                if (this._strMoveFolderStoreId == null)
                {
                    // Default
                    this._strMoveFolderStoreId = "";
                }
                return this._strMoveFolderStoreId;
            }
            set
            {
                this._strMoveFolderStoreId = value;
            }
        }

        /// <summary>
        /// The email subject of the report
        /// </summary>
        public string ReportSubject
        {
            get
            {
                if (this._strReportSubject == null)
                {
                    // Default
                    this._strReportSubject = "Spam Report";
                }
                return this._strReportSubject;
            }
            set
            {
                this._strReportSubject = value;
            }
        }

        /// <summary>
        /// Text to append to the end of each spam message in the report
        /// </summary>
        public string ReportEndText
        {
            get
            {
                if (this._strReportEndText == null)
                {
                    // Default
                    this._strReportEndText = "-- End of spam submission --";
                }
                return this._strReportEndText;
            }
            set
            {
                this._strReportEndText = value;
            }
        }

        /// <summary>
        /// Text to display in the report body
        /// </summary>
        public string MessageBody
        {
            get
            {
                if (this._strMessageBody == null)
                {
                    // Default
                    this._strMessageBody = "Please see the attached file for details";
                }
                return this._strMessageBody;
            }
            set
            {
                this._strMessageBody = value;
            }
        }

        /// <summary>
        /// An array of all the requested addresses to use in the To: field of the report
        /// </summary>
        public List<string> ToAddresses
        {
            get
            {
                if (this._arrToAddresses == null)
                {
                    this._arrToAddresses = new List<string>();
                }
                return this._arrToAddresses;
            }
        }

        /// <summary>
        /// An array of all the requested addresses to use in the BCC: field of the report
        /// </summary>
        public List<string> BccAddresses
        {
            get
            {
                if (this._arrBccAddresses == null)
                {
                    this._arrBccAddresses = new List<string>();
                }
                return this._arrBccAddresses;
            }
        }

        /// <summary>
        /// Whether or not to strip out headers that SpamCop cannot understand
        /// </summary>
        public bool CleanHeaders
        {
            get
            {
                if (this._blnCleanHeaders == null)
                {
                    // Default
                    this._blnCleanHeaders = false;
                }
                return (bool)this._blnCleanHeaders;
            }
            set
            {
                this._blnCleanHeaders = value;
            }
        }

        /// <summary>
        /// Whether or not to delete the spam after reporting it
        /// </summary>
        public bool DeleteAfterReport
        {
            get
            {
                if (this._blnDeleteAfterReport == null)
                {
                    // Default
                    this._blnDeleteAfterReport = true;
                }
                return (bool)this._blnDeleteAfterReport;
            }
            set
            {
                this._blnDeleteAfterReport = value;
            }
        }

        /// <summary>
        /// Whether or not to mark spam as read after reporting
        /// </summary>
        public bool MarkAsReadAfterReport
        {
            get
            {
                if (this._blnMarkAsReadAfterReport == null)
                {
                    // Default
                    this._blnMarkAsReadAfterReport = false;
                }
                return (bool)this._blnMarkAsReadAfterReport;
            }
            set
            {
                this._blnMarkAsReadAfterReport = value;
            }
        }

        /// <summary>
        /// Whether or not to move the spam to another folder
        /// </summary>
        public bool MoveToFolderAfterReport
        {
            get
            {
                if (this._blnMoveToFolderAfterReport == null)
                {
                    // Default
                    this._blnMoveToFolderAfterReport = false;
                }
                return (bool)this._blnMoveToFolderAfterReport;
            }
            set
            {
                this._blnMoveToFolderAfterReport = value;
            }
        }

        /// <summary>
        /// Whether or not to ask for confirmation before sending reports
        /// </summary>
        public bool AskVerify
        {
            get
            {
                if (this._blnAskVerify == null)
                {
                    // Default
                    this._blnAskVerify = false;
                }
                return (bool)this._blnAskVerify;
            }
            set
            {
                this._blnAskVerify = value;
            }
        }

        /// <summary>
        /// Whether or not to keep a copy of the report in the sent items folder
        /// </summary>
        public bool KeepCopy
        {
            get
            {
                if (this._blnKeepCopy == null)
                {
                    // Default
                    this._blnKeepCopy = true;
                }
                return (bool)this._blnKeepCopy;
            }
            set
            {
                this._blnKeepCopy = value;
            }
        }

        /// <summary>
        /// Whether or not to send all spam messages as one report. If set to
        /// false then one report email will be used per spam
        /// </summary>
        public bool SendMultiple
        {
            get
            {
                if (this._blnSendMultiple == null)
                {
                    // Default
                    this._blnSendMultiple = false;
                }
                return (bool)this._blnSendMultiple;
            }
            set
            {
                this._blnSendMultiple = value;
            }
        }

        /// <summary>
        /// Whether or not to use RFC822 to attach the emails. Does not work with Spamcop.
        /// </summary>
        public bool UseRFC
        {
            get
            {
                if (this._blnUseRFC == null)
                {
                    // Default
                    this._blnUseRFC = false;
                }
                return (bool)this._blnUseRFC;
            }
            set
            {
                this._blnUseRFC = value;
            }
        }

        /// <summary>
        /// Attempt a send / receive after reporting
        /// </summary>
        public bool SendReceiveAfterReport
        {
            get
            {
                if (this._blnSendReceiveAfterReport == null)
                {
                    // Default
                    this._blnSendReceiveAfterReport = true;
                }
                return (bool)this._blnSendReceiveAfterReport;
            }
            set
            {
                this._blnSendReceiveAfterReport = value;
            }
        }

        /// <summary>
        /// The user-defined friendly name of the profile
        /// </summary>
        public string Name
        {
            get
            {
                if (this._strProfileName == null)
                {
                    // Default
                    this._strProfileName = "";
                }
                return this._strProfileName;
            }
            set
            {
                this._strProfileName = value;
            }
        }

        /// <summary>
        /// The GUID of the profile as stored in the registry
        /// </summary>
        public string Id
        {
            get
            {
                return this._strProfileId;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Loads the specified profile from the registry
        /// </summary>
        private void LoadProfile()
        {
            string strKey = SGGlobals.BaseRegistryKey + "Profiles\\" + this._strProfileId;
            
            // Make sure the key exists
            if (!SGGlobals.DoesKeyExist(strKey))
            {
                throw new ProfileNotFoundException("The profile " + this._strProfileId +
                    " Could not be found");
            }

            // Load the values into the class data
            this._strProfileName = (string)SGGlobals.LoadValue(strKey, "ProfileName", "Default");
            this._blnDeleteAfterReport = (bool)SGGlobals.LoadValue(strKey, "DeleteAfterReport", false);
            this._blnMarkAsReadAfterReport = (bool)SGGlobals.LoadValue(strKey, "MarkAsReadAfterReport", false);
            this._blnMoveToFolderAfterReport = (bool)SGGlobals.LoadValue(strKey, "MoveToFolderAfterReport", false);
            this._strMoveFolderName = (string)SGGlobals.LoadValue(strKey, "MoveFolderName", "");
            this._strMoveFolderStoreId = (string)SGGlobals.LoadValue(strKey, "MoveFolderStoreId", "");
            this._strReportSubject = (string)SGGlobals.LoadValue(strKey, "ReportSubject", "Spam Report");
            this._strMessageBody = (string)SGGlobals.LoadValue(strKey, "MessageBody", "");
            this._strReportEndText = (string)SGGlobals.LoadValue(strKey, "ReportEndText", "");
            this._blnAskVerify = (bool)SGGlobals.LoadValue(strKey, "AskVerify", true);
            this._blnKeepCopy = (bool)SGGlobals.LoadValue(strKey, "KeepCopy", true);
            this._blnSendMultiple = (bool)SGGlobals.LoadValue(strKey, "SendMultiple", true);
            this._blnUseRFC = (bool)SGGlobals.LoadValue(strKey, "UseRFC", false);
            this._blnSendReceiveAfterReport = (bool)SGGlobals.LoadValue(strKey, "SendReceiveAfterReport", false);
            this._blnCleanHeaders = (bool)SGGlobals.LoadValue(strKey, "CleanHeaders", false);

            // Load the addresses into the array lists
            char[] splitter = { ';' };
            string[] strToAddresses = ((string)SGGlobals.LoadValue(strKey, "ToAddress", "")).Split(splitter);
            this._arrToAddresses = new List<string>();
            foreach (string strAddress in strToAddresses)
            {
                if (!string.IsNullOrEmpty(strAddress)) { this._arrToAddresses.Add(strAddress); }
            }
            string[] strBccAddresses = ((string)SGGlobals.LoadValue(strKey, "BccAddress", "")).Split(splitter);
            this._arrBccAddresses = new List<string>();
            foreach (string strAddress in strBccAddresses)
            {
                if (!string.IsNullOrEmpty(strAddress)) { this._arrBccAddresses.Add(strAddress); }
            }
        }

        /// <summary>
        /// Saves the current profile to the registry
        /// </summary>
        public void Save()
        {
            // Declare registry key we will be working with
            //RegistryKey regProfileKey;
            string strKey = SGGlobals.BaseRegistryKey + "Profiles\\" + this._strProfileId;

            // Determine if this is a new profile
            if (SGGlobals.DoesKeyExist(strKey) == false)
            {
                // Create the new registry key
                //regProfileKey = Registry.CurrentUser.CreateSubKey(SGGlobals.BaseRegistryKey +
                //    "Profiles\\" + this._strProfileId);
                SGGlobals.CreateKey(this._strProfileId, SGGlobals.BaseRegistryKey + "Profiles");
            }
            //else
            //{
            //    // Open the existing key
            //    regProfileKey = Registry.CurrentUser.OpenSubKey(SGGlobals.BaseRegistryKey +
            //        "Profiles\\" + this._strProfileId, true);
            //}

            // Set all the saved standard information
            //regProfileKey.SetValue("ProfileName", this.Name, RegistryValueKind.String);
            //regProfileKey.SetValue("CleanHeaders", this.CleanHeaders, RegistryValueKind.DWord);
            //regProfileKey.SetValue("DeleteAfterReport", this.DeleteAfterReport, RegistryValueKind.DWord);
            //regProfileKey.SetValue("MarkAsReadAfterReport", this.MarkAsReadAfterReport, RegistryValueKind.DWord);
            //regProfileKey.SetValue("MoveToFolderAfterReport", this.MoveToFolderAfterReport, RegistryValueKind.DWord);
            //regProfileKey.SetValue("MoveFolderName", this.MoveFolderName, RegistryValueKind.String);
            //regProfileKey.SetValue("MoveFolderStoreId", this.MoveFolderStoreId, RegistryValueKind.String);
            //regProfileKey.SetValue("ReportSubject", this.ReportSubject, RegistryValueKind.String);
            //regProfileKey.SetValue("ReportEndText", this.ReportEndText, RegistryValueKind.String);
            //regProfileKey.SetValue("MessageBody", this.MessageBody, RegistryValueKind.String);
            //regProfileKey.SetValue("AskVerify", this.AskVerify, RegistryValueKind.DWord);
            //regProfileKey.SetValue("KeepCopy", this.KeepCopy, RegistryValueKind.DWord);
            //regProfileKey.SetValue("SendMultiple", this.SendMultiple, RegistryValueKind.DWord);
            //regProfileKey.SetValue("IncludeIPAddress", this.IncludeIPAddress, RegistryValueKind.DWord);
            //regProfileKey.SetValue("UseRFC", this.UseRFC, RegistryValueKind.DWord);
            //regProfileKey.SetValue("SendReceiveAfterReport", this.SendReceiveAfterReport, RegistryValueKind.DWord);
            //regProfileKey.SetValue("FixMIME", this.FixMIME, RegistryValueKind.DWord);
            //regProfileKey.SetValue("ReportToAllProfiles", this.ReportToAllProfiles, RegistryValueKind.DWord);
            //regProfileKey.SetValue("ProfileType", this.ProfileType, RegistryValueKind.DWord);

            SGGlobals.SaveSetting(strKey,"ProfileName", this.Name);
            SGGlobals.SaveSetting(strKey, "CleanHeaders", this.CleanHeaders);
            SGGlobals.SaveSetting(strKey, "DeleteAfterReport", this.DeleteAfterReport);
            SGGlobals.SaveSetting(strKey, "MarkAsReadAfterReport", this.MarkAsReadAfterReport);
            SGGlobals.SaveSetting(strKey, "MoveToFolderAfterReport", this.MoveToFolderAfterReport);
            SGGlobals.SaveSetting(strKey, "MoveFolderName", this.MoveFolderName);
            SGGlobals.SaveSetting(strKey, "MoveFolderStoreId", this.MoveFolderStoreId);
            SGGlobals.SaveSetting(strKey, "ReportSubject", this.ReportSubject);
            SGGlobals.SaveSetting(strKey, "ReportEndText", this.ReportEndText);
            SGGlobals.SaveSetting(strKey, "MessageBody", this.MessageBody);
            SGGlobals.SaveSetting(strKey, "AskVerify", this.AskVerify);
            SGGlobals.SaveSetting(strKey, "KeepCopy", this.KeepCopy);
            SGGlobals.SaveSetting(strKey, "SendMultiple", this.SendMultiple);
            SGGlobals.SaveSetting(strKey, "UseRFC", this.UseRFC);
            SGGlobals.SaveSetting(strKey, "SendReceiveAfterReport", this.SendReceiveAfterReport);
            // Save the array lists
            string strToAddress = "";
            string strBccAddress = "";
            foreach (string strAddress in this.ToAddresses)
            {
                // remove excess ;
                if (strAddress != string.Empty)
                    strToAddress += strAddress.Replace(";","") + ";";
            }
            foreach (string strAddress in this.BccAddresses)
            {
                if (strAddress!=string.Empty)
                    strBccAddress += strAddress.Replace(";", "") + ";";
            }
            SGGlobals.SaveSetting(strKey, "ToAddress", strToAddress);
            SGGlobals.SaveSetting(strKey, "BccAddress", strBccAddress);
        }

        /// <summary>
        /// Check if the current profile already exists in the registry
        /// </summary>
        /// <returns>True if profile exists</returns>
        //private bool ProfileExists()
        //{
        //    // Get the root HKCU key
        //    RegistryKey regSettings = Registry.CurrentUser;

        //    // Get the selected profile
        //    regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey +
        //        "Profiles\\" + this._strProfileId);

        //    // Does the key exist?
        //    if (regSettings == null)
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}

        #endregion
    }

    #endregion

    #region Exception Class

    /// <summary>
    /// Exception thrown if calling code attempts to load
    /// a profile that does not exist
    /// </summary>
    public class ProfileNotFoundException : Exception
    {
        public ProfileNotFoundException() : base() { }
        public ProfileNotFoundException(string message) : base(message) { }
    }

    #endregion

    #region Profile list class

    /// <summary>
    /// Static class to return a list of the current user's profile IDs
    /// </summary>
    public static class UserProfiles
    {
        private static ArrayList _arrProfiles;

        /// <summary>
        /// Returns an array list of all the user's profile IDs
        /// </summary>
        public static ArrayList ProfileList
        {
            get
            {
                if (_arrProfiles == null)
                {
                    // Make sure the root key exists
                    SGGlobals.CreateBaseKey();

                    // Populate the list
                    _arrProfiles = new ArrayList();

                    // Get the root HKCU key
                    RegistryKey regSettings = Registry.CurrentUser;

                    // Get all the subkeys
                    regSettings = regSettings.OpenSubKey(SGGlobals.BaseRegistryKey +
                        "Profiles\\");
                    string[] strProfiles = regSettings.GetSubKeyNames();
                    foreach (string strProfile in strProfiles)
                    {
                        _arrProfiles.Add(new Profile(strProfile));
                    }
                }
                // Return the value
                return _arrProfiles;
            }
        }

        /// <summary>
        /// Clears the cache of user profiles to allow the calling
        /// code to load in the latest settings
        /// </summary>
        public static void ClearProfileCache()
        {
            _arrProfiles = null;
        }

        /// <summary>
        /// Get a profile from the collection by name
        /// </summary>
        /// <param name="pstrName"></param>
        /// <returns></returns>
        public static Profile GetProfileByName(string pstrName)
        {
            Profile objProfile = null;

            // Loop through all the profiles
            foreach (Profile objTempProfile in ProfileList)
            {
                if (objTempProfile.Name == pstrName)
                {
                    // Found the correct one, set it to be the return profile
                    objProfile = objTempProfile;
                }
            }
            if (objProfile == null)
            {
                throw new ProfileNotFoundException("The selected profile name does not exist");
            }

            return objProfile;
        }

        /// <summary>
        /// Deletes a profile from the registry
        /// </summary>
        /// <param name="pstrProfileName"></param>
        public static void DeleteProfile(string pstrProfileName)
        {
            // Get the ID of the profile (throws exception if not found)
            string strProfileId = GetProfileByName(pstrProfileName).Id;

            // Get the root HKCU key
            RegistryKey regSettings = Registry.CurrentUser.OpenSubKey(
                SGGlobals.BaseRegistryKey + "Profiles\\", true);

            // Delete the key
            regSettings.DeleteSubKeyTree(strProfileId);
        }
    }

    #endregion
}
