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
    public partial class ctlManageProfiles : UserControl
    {
        #region Class Data

        public delegate void EditProfileRaisedHandler(object sender);
        public delegate void ProfileListChangedHandler(object sender);
        public event EditProfileRaisedHandler EditProfileRaised;
        public event ProfileListChangedHandler ProfileListChanged;

        private string _strSelectedProfile;
        private string _strDefaultSpamText = " (spam default)";
        private string _strDefaultHamText = " (ham default)";
        private string _pNameExtensionSpam = " (spam)";
        private string _pNameExtentionHam = " (ham)";

        #endregion

        #region Constructor

        public ctlManageProfiles()
        {
            InitializeComponent();
            PopulateList();
        }

        #endregion

        #region Properties

        public string SelectedProfile
        {
            get
            {
                return this._strSelectedProfile;
            }
        }

        #endregion

        #region Methods

        public void PopulateList()
        {
            lbProfiles.Items.Clear();

            foreach (Profile objProfile in UserProfiles.ProfileList)
            {
                string pNameExtension = string.Empty;

                if (objProfile.ProfileType == Profile.eProfileType.SpamProfile)
                {
                    pNameExtension = _pNameExtensionSpam;
                }
                else if (objProfile.ProfileType == Profile.eProfileType.HamProfile)
                {
                    pNameExtension = _pNameExtentionHam;
                }
                else
                {
                    pNameExtension = _pNameExtensionSpam + _pNameExtentionHam;
                }


                if (GlobalSettings.DefaultSpamProfileId.Equals(objProfile.Id))
                {
                    // See if this is also the default ham (??)
                    if (GlobalSettings.DefaultHamProfileId.Equals(objProfile.Id))
                    {
                        lbProfiles.Items.Add(objProfile.Name + pNameExtension + _strDefaultSpamText + _strDefaultHamText);
                    }
                    else
                    {
                        lbProfiles.Items.Add(objProfile.Name + pNameExtension + _strDefaultSpamText);
                    }
                }
                else if (GlobalSettings.DefaultHamProfileId.Equals(objProfile.Id))
                {
                    lbProfiles.Items.Add(objProfile.Name + pNameExtension + _strDefaultHamText);
                }
                else
                {
                    lbProfiles.Items.Add(objProfile.Name + pNameExtension);
                }


            }

            if (lbProfiles.SelectedIndex > -1)
            {
                SetButtonStates(true);
            }
            else
            {
                SetButtonStates(false);
            }
        }

        private void SetButtonStates(bool pblnEnableButtons)
        {
            this.btnDeleteProfile.Enabled = pblnEnableButtons;
            this.btnEditProfile.Enabled = pblnEnableButtons;
            this.btnSetAsHamProfile.Enabled = pblnEnableButtons;
            this.btnSetAsSpamProfile.Enabled = pblnEnableButtons;

            // Also check this profile isn't currently either of the defaults
            if (lbProfiles.SelectedIndex > -1 &&
                GlobalSettings.DefaultSpamProfileId.Equals(UserProfiles.GetProfileByName(GetSelectedProfileName()).Id) &&
                GlobalSettings.DefaultHamProfileId.Equals(UserProfiles.GetProfileByName(GetSelectedProfileName()).Id))
            {
                this.btnSetSpamDefault.Enabled = false;
                this.btnSetHamDefault.Enabled = false;
            }
            else if (lbProfiles.SelectedIndex > -1 &&
                GlobalSettings.DefaultSpamProfileId.Equals(UserProfiles.GetProfileByName(GetSelectedProfileName()).Id))
            {
                this.btnSetSpamDefault.Enabled = false;
                this.btnSetHamDefault.Enabled = pblnEnableButtons;
            }
            else if (lbProfiles.SelectedIndex > -1 &&
                GlobalSettings.DefaultHamProfileId.Equals(UserProfiles.GetProfileByName(GetSelectedProfileName()).Id))
            {
                this.btnSetHamDefault.Enabled = false;
                this.btnSetSpamDefault.Enabled = pblnEnableButtons;
            }
            else
            {
                this.btnSetSpamDefault.Enabled = pblnEnableButtons;
                this.btnSetHamDefault.Enabled = pblnEnableButtons;
            }
        }

        /// <summary>
        /// Helper function to get the name of the selected profile
        /// </summary>
        /// <returns></returns>
        private string GetSelectedProfileName()
        {
            if (lbProfiles.SelectedIndex > -1)
            {
                string strItemText = lbProfiles.SelectedItem.ToString();

                // Worst case scenario, both default?
                /* if (strItemText.Length > (_strDefaultSpamText.Length + _strDefaultHamText.Length) &&
                     strItemText.Substring(strItemText.Length - (_strDefaultSpamText.Length + _strDefaultHamText.Length),
                     (_strDefaultSpamText.Length + _strDefaultHamText.Length)).Equals(_strDefaultSpamText + _strDefaultHamText))
                 {
                     return strItemText.Substring(0, strItemText.Length - (_strDefaultSpamText.Length + _strDefaultHamText.Length));
                 }
                 // Spam default
                 else if (strItemText.Length > _strDefaultSpamText.Length &&
                     strItemText.Substring(strItemText.Length - _strDefaultSpamText.Length,
                     _strDefaultSpamText.Length).Equals(_strDefaultSpamText))
                 {
                     return strItemText.Substring(0, strItemText.Length - _strDefaultSpamText.Length);
                 }
                 // Ham default
                 else if (strItemText.Length > _strDefaultHamText.Length &&
                     strItemText.Substring(strItemText.Length - _strDefaultHamText.Length,
                     _strDefaultHamText.Length).Equals(_strDefaultHamText))
                 {
                     return strItemText.Substring(0, strItemText.Length - _strDefaultHamText.Length);
                 }
                 else
                 {
                 */
                // process the profile name extensions
                // Normal
                // Delete the extension text.
                strItemText = strItemText.Replace(_pNameExtensionSpam, "");
                strItemText = strItemText.Replace(_pNameExtentionHam, "");
                strItemText = strItemText.Replace(_strDefaultHamText, "");
                strItemText = strItemText.Replace(_strDefaultSpamText, "");
                return strItemText;
                //}
            }
            else
            {
                return "";
            }
        }

        #endregion

        #region Event Handlers

        private void btnAddProfile_Click(object sender, EventArgs e)
        {
            frmNewProfile objNewProfile = new frmNewProfile();
            objNewProfile.ShowDialog();

            if (objNewProfile.Result == DialogResult.OK)
            {
                // Clear all required caches and fire update event
                UserProfiles.ClearProfileCache();
                PopulateList();
                ProfileListChanged(this);
            }
        }

        private void btnEditProfile_Click(object sender, EventArgs e)
        {
            this._strSelectedProfile = GetSelectedProfileName();
            EditProfileRaised(this);
        }

        private void btnDeleteProfile_Click(object sender, EventArgs e)
        {
            if (lbProfiles.SelectedIndex > -1)
            {
                // Make sure they are sure!
                if (MessageBox.Show("Are you sure you want to delete this profile?" +
                    Environment.NewLine + "This cannot be undone!",
                    "SpamGrabber", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // See if this was one of the defaults
                    Profile objProfile = UserProfiles.GetProfileByName(GetSelectedProfileName());
                    if (GlobalSettings.DefaultSpamProfileId.Equals(objProfile.Id))
                    {
                        GlobalSettings.DefaultSpamProfileId = string.Empty;
                        GlobalSettings.ResetDefaultProfile(GlobalSettings.DefaultType.Spam);
                    }
                    if (GlobalSettings.DefaultHamProfileId.Equals(objProfile.Id))
                    {
                        GlobalSettings.DefaultHamProfileId = string.Empty;
                        GlobalSettings.ResetDefaultProfile(GlobalSettings.DefaultType.Ham);
                    }

                    // Delete the profile
                    UserProfiles.DeleteProfile(objProfile.Name);

                    // Clear all required caches and fire update event
                    UserProfiles.ClearProfileCache();
                    PopulateList();
                    ProfileListChanged(this);
                }
            }
        }

        private void lbProfiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbProfiles.SelectedIndex > -1)
            {
                SetButtonStates(true);
            }
            else
            {
                SetButtonStates(false);
            }
        }

        private void lbProfiles_Leave(object sender, EventArgs e)
        {
            if (lbProfiles.SelectedIndex > -1)
            {
                SetButtonStates(true);
            }
            else
            {
                SetButtonStates(false);
            }
        }

        private void lbProfiles_DoubleClick(object sender, EventArgs e)
        {
            this._strSelectedProfile = GetSelectedProfileName();
            EditProfileRaised(this);
        }

        /// <summary>
        /// Sets the selected profile to be the default Spam reporting address
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetDefault_Click(object sender, EventArgs e)
        {
            Profile theProfile = UserProfiles.GetProfileByName(GetSelectedProfileName());
            // Set the global setting
            GlobalSettings.DefaultSpamProfileId = theProfile.Id;

            if (theProfile.ProfileType == Profile.eProfileType.HamProfile)
                theProfile.ProfileType = Profile.eProfileType.SpamHamProfile;
            else
                theProfile.ProfileType = Profile.eProfileType.SpamProfile;

            // Reload the listbox
            PopulateList();
        }

        /// <summary>
        /// Sets the selected profile to be the default Ham reporting address
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetHamDefault_Click(object sender, EventArgs e)
        {
            Profile theProfile = UserProfiles.GetProfileByName(GetSelectedProfileName());
            // Set the global setting
            GlobalSettings.DefaultHamProfileId = theProfile.Id;

            if (theProfile.ProfileType == Profile.eProfileType.SpamProfile)
                theProfile.ProfileType = Profile.eProfileType.SpamHamProfile;
            else
                theProfile.ProfileType = Profile.eProfileType.HamProfile;
            // Reload the listbox
            PopulateList();
        }

        private void btnSetAsSpamProfile_Click(object sender, EventArgs e)
        {
            Profile theProfile = UserProfiles.GetProfileByName(GetSelectedProfileName());

            if (theProfile.ProfileType == Profile.eProfileType.HamProfile)
                theProfile.ProfileType = Profile.eProfileType.SpamHamProfile;
            else if (theProfile.ProfileType == Profile.eProfileType.SpamHamProfile)
            {
                theProfile.ProfileType = Profile.eProfileType.HamProfile;

                if (theProfile.Id == GlobalSettings.DefaultSpamProfileId)
                {
                    GlobalSettings.DefaultSpamProfileId = "";
                    GlobalSettings.Save();
                }
            }
            else
                theProfile.ProfileType = Profile.eProfileType.SpamProfile;
            // Reload the listbox
            PopulateList();
        }

        private void btnSetAsHamProfile_Click(object sender, EventArgs e)
        {
            Profile theProfile = UserProfiles.GetProfileByName(GetSelectedProfileName());

            if (theProfile.ProfileType == Profile.eProfileType.SpamProfile)
                theProfile.ProfileType = Profile.eProfileType.SpamHamProfile;
            else if (theProfile.ProfileType == Profile.eProfileType.SpamHamProfile)
            {
                theProfile.ProfileType = Profile.eProfileType.SpamProfile;
                if (theProfile.Id == GlobalSettings.DefaultHamProfileId)
                {
                    GlobalSettings.DefaultHamProfileId = "";
                    GlobalSettings.Save();
                }
            }
            else
                theProfile.ProfileType = Profile.eProfileType.HamProfile;
            // Reload the listbox
            PopulateList();

        }


        private void ctlManageProfiles_Load(object sender, EventArgs e)
        {
            ToolTip ttManageProfiles = new ToolTip();

            // Set up the delays for the ToolTip.
            ttManageProfiles.AutoPopDelay = 5000;
            ttManageProfiles.InitialDelay = 1000;
            ttManageProfiles.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            ttManageProfiles.ShowAlways = true;

            ttManageProfiles.SetToolTip(btnAddProfile, "Adds a new profile to the configuration.");
            ttManageProfiles.SetToolTip(btnDeleteProfile, "Deletes the selected profile from the configuration.");
            ttManageProfiles.SetToolTip(btnEditProfile, "Edits the selected profile.");
            ttManageProfiles.SetToolTip(btnSetAsHamProfile, "Sets the selected profile as a HAM profile.");
            ttManageProfiles.SetToolTip(btnSetAsSpamProfile, "Sets the selected profile as a SPAM profile.");
            ttManageProfiles.SetToolTip(btnSetHamDefault, "Sets the selected profile as the default HAM profile.");
            ttManageProfiles.SetToolTip(btnSetSpamDefault, "Sets the selected profile as the default SPAM profile.");
        }
        #endregion
    }
}
