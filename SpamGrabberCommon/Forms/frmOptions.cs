using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SpamGrabberCommon;
using System.Collections;

namespace SpamGrabberControl
{
    public partial class frmOptions : Form
    {
        //Implemented opacity on the for load event.
        #region Class Data

        private ArrayList _arrLoadedProfiles;
        Timer _tmrFade;

        #endregion

        #region Constructors

        public frmOptions()
        {
            InitializeComponent();
            ManageProfilesControl.EditProfileRaised += new ctlManageProfiles.EditProfileRaisedHandler(ManageProfilesControl_EditProfileRaised);
            ManageProfilesControl.ProfileListChanged += new ctlManageProfiles.ProfileListChangedHandler(ManageProfilesControl_ProfileListChanged);
            _tmrFade = new Timer();
            _tmrFade.Tick += new EventHandler(tmrFade_Tick);
            _tmrFade.Interval = 10;
            this.Opacity = 0;// make the form transparent
        }

        void tmrFade_Tick(object sender, EventArgs e)
        {
            this.Opacity += 0.05;
            if (this.Opacity >= .95)
            {
                this.Opacity = 1;
                _tmrFade.Enabled = false;
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// The profiles that have been loaded and therefore require saving
        /// on closing the form
        /// </summary>
        private ArrayList LoadedProfiles
        {
            get
            {
                if (this._arrLoadedProfiles == null)
                {
                    this._arrLoadedProfiles = new ArrayList();
                }
                return this._arrLoadedProfiles;
            }
        }

        #endregion

        #region Methods

        private void LoadTreeView()
        {
            // Clear all existing profiles
            tvwOptions.Nodes["ndeProfiles"].Nodes.Clear();

            // Get a reference to the profiles node
            TreeNode ndeProfiles = tvwOptions.Nodes["ndeProfiles"];

            // Get all the user's profiles
            foreach (Profile objProfile in UserProfiles.ProfileList)
            {
                ndeProfiles.Nodes.Add(new TreeNode(objProfile.Name));
            }
        }

        private void CacheLoadedProfile()
        {
            // If the profile control had a profile loaded into it then
            // put the profile into our collection to save / discard on closing
            if (submissionControl.Profile != null)
            {
                foreach (Profile objProfile in this.LoadedProfiles)
                {
                    if (objProfile.Id == submissionControl.Profile.Id)
                    {
                        // Update the profile
                        this.LoadedProfiles.Remove(objProfile);
                        break;
                    }
                }
                // Add the profile to the collection
                this.LoadedProfiles.Add(submissionControl.Profile);
            }
        }

        #endregion

        #region Event Handlers

        private void tvwOptions_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //we need to disable screenupdating while processing the configuration.
            UIControl.StopWindowUpdating(this.Handle);
            // Cache profile, if one has been selected
            CacheLoadedProfile();

            tvwOptions.Visible = true;

            TreeNode selNode = e.Node;
            if (selNode.Parent != null && selNode.Parent.Name == "ndeProfiles")
            {
                // This is a profile node. Load the profile
                //
                // Show / hide the controls
                generalOptions.Visible = false;
                about.Visible = false;
                submissionControl.Visible = true;
                ManageProfilesControl.Visible = false;

                // Populate the settings control
                submissionControl.Profile = UserProfiles.GetProfileByName(selNode.Text);

                // Cache profile
                CacheLoadedProfile();
            }
            else
            {

                switch (selNode.Name)
                {
                    case "ndeGlobal":
                        generalOptions.Visible = true;
                        about.Visible = false;
                        submissionControl.Visible = false;
                        ManageProfilesControl.Visible = false;
                        break;
                    case "ndeAbout":
                        generalOptions.Visible = false;
                        about.Visible = true;
                        submissionControl.Visible = false;
                        ManageProfilesControl.Visible = false;
                        break;
                    case "ndeProfiles":
                        generalOptions.Visible = false;
                        about.Visible = false;
                        submissionControl.Visible = false;
                        ManageProfilesControl.Visible = true;
                        break;
                    default:
                        generalOptions.Visible = false;
                        about.Visible = true;
                        submissionControl.Visible = false;
                        ManageProfilesControl.Visible = false;
                        break;
                }
            }

            UIControl.StartWindowUpdating(this.Handle);
            this.Invalidate(true);

        }

        private void frmOptions_Load(object sender, EventArgs e)
        {
            //UIControl.StopWindowUpdating(this.Handle);
            LoadTreeView();
            tvwOptions.SelectedNode = tvwOptions.Nodes[0];
            generalOptions.Visible = true;
            //generalOptions.Invalidate(true);
            //UIControl.StartWindowUpdating(this.Handle);
            this.Invalidate(true);
            tvwOptions.Focus();
            _tmrFade.Enabled = true;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            // Clear the profile cache
            UserProfiles.ClearProfileCache();

            // Clear the loaded profile cache
            this.LoadedProfiles.Clear();

            // Reset any changes made to global settings
            GlobalSettings.LoadSettings();

            // Close the form
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            // Save the profiles in our collection
            foreach (Profile objProfile in this.LoadedProfiles)
            {
                objProfile.Save();
            }

            // Save the global settings
            GlobalSettings.Save();

            // Close the form
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void ManageProfilesControl_EditProfileRaised(object sender)
        {
            // Show / hide the controls
            UIControl.StopWindowUpdating(this.Handle);
            generalOptions.Visible = false;
            about.Visible = false;
            submissionControl.Visible = true;
            ManageProfilesControl.Visible = false;

            // Populate the settings control
            submissionControl.Profile = UserProfiles.GetProfileByName(ManageProfilesControl.SelectedProfile);

            // Select the item in the treeview
            TreeNode objSelectedProfile = null;
            foreach (TreeNode objNode in tvwOptions.Nodes["ndeProfiles"].Nodes)
            {
                if (objNode.Text.Equals(ManageProfilesControl.SelectedProfile))
                {
                    objSelectedProfile = objNode;
                }
            }
            if (objSelectedProfile != null)
            {
                tvwOptions.SelectedNode = objSelectedProfile;
            }

            // Cache the loaded profile
            CacheLoadedProfile();
            UIControl.StartWindowUpdating(this.Handle);
        }

        private void ManageProfilesControl_ProfileListChanged(object sender)
        {
            // Clear the profile cache just in case
            UserProfiles.ClearProfileCache();

            // Update the treeview
            LoadTreeView();
        }

        #endregion

        private void about_Load(object sender, EventArgs e)
        {

        }


    }
}