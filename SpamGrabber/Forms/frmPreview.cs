#region Imports

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

#endregion

namespace SpamGrabberControl
{
    /// <summary>
    /// Message preview form
    /// </summary>
    public partial class frmPreview : Form
    {
        #region Class Data

        private ArrayList _arrItems = null;
        private string _strDialogTitle = "Safe Preview of '{0}'";
        private int _intCurrentIndex;

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public frmPreview()
        {
            InitializeComponent();
        }

        #endregion

        #region Properies

        /// <summary>
        /// Collection of MailItems to display
        /// </summary>
        public ArrayList Items
        {
            get
            {
                if (this._arrItems == null)
                {
                    this._arrItems = new ArrayList();
                }
                return this._arrItems;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Clears the item list
        /// </summary>
        public void ClearItems()
        {
            this._arrItems = new ArrayList();
        }

        /// <summary>
        /// Displays the form. Items() must be populated beforehand.
        /// </summary>
        /// <returns></returns>
        public new DialogResult ShowDialog()
        {
            // Disable the buttons by default
            btnPrevious.Enabled = false;
            btnNext.Enabled = false;

            // If there are no items to display, close the form
            if (Items.Count == 0)
            {
                this.Close();
                return DialogResult.Abort;
            }
            // If we have more than one email, show the next button
            else if (Items.Count > 1)
            {
                btnNext.Enabled = true;
            }

            // Set the index and show the first item
            _intCurrentIndex = 0;
            ShowItem(_intCurrentIndex);

            // Show the dialog
            base.ShowDialog();

            return DialogResult.OK;
        }

        /// <summary>
        /// Displays the selected email
        /// </summary>
        /// <param name="pintIndex"></param>
        private void ShowItem(int pintIndex)
        {
            if (pintIndex >= Items.Count)
            {
                return;
            }
            string strSubject = string.Empty;
            string strBody = string.Empty;

            if (Items[pintIndex] is Outlook.MailItem)
            {
                // Get the item
                Outlook.MailItem objItem = (Outlook.MailItem)Items[pintIndex];

                // Extract info
                strSubject = objItem.Subject;
                strBody = objItem.Body;
            }
            else if (Items[pintIndex] is Outlook.PostItem)
            {
                Outlook.PostItem objItem = (Outlook.PostItem)Items[pintIndex];

                // Extract info
                strSubject = objItem.Subject;
                strBody = objItem.Body;
            }

            // Display info
            this.Text = string.Format(_strDialogTitle, strSubject);
            this.txtBody.Text = strBody;
            this.txtBody.SelectionStart = 0;
            this.txtBody.SelectionLength = 0;
        }

        #endregion

        #region Event Handlers

        /// <summary>
        /// Handles the Previous button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPrevious_Click(object sender, EventArgs e)
        {
            // Decrement counter
            _intCurrentIndex--;

            // Enable / disable buttons based on count
            if (_intCurrentIndex == 0)
            {
                this.btnPrevious.Enabled = false;
            }
            else
            {
                this.btnPrevious.Enabled = true;
            }
            this.btnNext.Enabled = true;

            // Display email
            ShowItem(_intCurrentIndex);
        }

        /// <summary>
        /// Handles the Next button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNext_Click(object sender, EventArgs e)
        {
            // Increment counter
            _intCurrentIndex++;

            // Enable / disable buttons based on count
            if (_intCurrentIndex == Items.Count - 1)
            {
                this.btnNext.Enabled = false;
            }
            else
            {
                this.btnNext.Enabled = true;
            }
            this.btnPrevious.Enabled = true;

            // Display email
            ShowItem(_intCurrentIndex);
        }

        /// <summary>
        /// Closes the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}