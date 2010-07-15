#region Imports

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SpamGrabberCommon;

#endregion

namespace SpamGrabberControl
{
    /// <summary>
    /// Prompts the user to select a profile
    /// to report to
    /// </summary>
    public partial class frmSelectProfile : Form
    {
        #region Class Data

        private Profile _objSelectedProfile;
        SGGlobals.ReportAction _enumAction;

        #endregion

        #region Properties

        /// <summary>
        /// The profile selected by the user
        /// </summary>
        public Profile SelectedProfile
        {
            get
            {
                return this._objSelectedProfile;
            }
        }

        /// <summary>
        /// The selected reporting action
        /// </summary>
        public SGGlobals.ReportAction ReportAction
        {
            get
            {
                return this._enumAction;
            }
        }

        #endregion

        #region Constructor

        public frmSelectProfile()
        {
            InitializeComponent();
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// Populates the selected profile and closes the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, EventArgs e)
        {
            // Make sure they have selected a profile
            if (cboProfile.Text.Equals(""))
            {
                MessageBox.Show("Please select a valid profile", "Select Profile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Set the selected profile
            this._objSelectedProfile = UserProfiles.GetProfileByName(cboProfile.Text);

            // Set the reporting action
            if (radSpam.Checked == true)
            {
                this._enumAction = SGGlobals.ReportAction.ReportSpam;
            }
            else
            {
                this._enumAction = SGGlobals.ReportAction.ReportHam;
            }

            // Set dialog result
            this.DialogResult = DialogResult.OK;

            // Close the form
            this.Close();
        }

        /// <summary>
        /// Closes the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// Load event for the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmSelectProfile_Load(object sender, EventArgs e)
        {
            // Clear the combobox
            cboProfile.Items.Clear();

            // Add all the profiles to the combobox
            foreach (Profile objProfile in UserProfiles.ProfileList)
            {
                cboProfile.Items.Add(objProfile.Name);
            }
        }

        #endregion
    }
}