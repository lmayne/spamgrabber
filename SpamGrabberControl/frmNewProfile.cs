using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SpamGrabberCommon;

namespace SpamGrabberControl
{
    public partial class frmNewProfile : Form
    {

        private DialogResult _objDialogStatus;

        public DialogResult Result
        {
            get
            {
                return _objDialogStatus;
            }
        }

        public frmNewProfile()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this._objDialogStatus = DialogResult.Cancel;
            this.Close();
        }

        private void btnAddProfile_Click(object sender, EventArgs e)
        {
            // Make sure the email is valid
            if (SGGlobals.IsEmailValid(txtDefaultAddress.Text) == false)
            {
                MessageBox.Show("Email address does not appear to be valid!", "Error creating profile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Make sure the profile doesn't exist
            try
            {
                UserProfiles.GetProfileByName(txtProfileName.Text);

                // Profile exists!
                MessageBox.Show("Profile name already exists!", "Error creating profile", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            catch (ProfileNotFoundException)
            {
                // Profile does not exist, create it
                Profile objNewProfile = new Profile();
                objNewProfile.Name = txtProfileName.Text;
                objNewProfile.ToAddresses.Add(txtDefaultAddress.Text);
                objNewProfile.Save();
                this._objDialogStatus = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An unknown error occured creating the new profile:"
                    + Environment.NewLine + ex.Message, "Error creating profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}