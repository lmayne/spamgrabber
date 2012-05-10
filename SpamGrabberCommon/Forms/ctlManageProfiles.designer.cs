namespace SpamGrabberControl
{
    partial class ctlManageProfiles
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.grpProfiles = new System.Windows.Forms.GroupBox();
            this.btnSetHamDefault = new System.Windows.Forms.Button();
            this.btnSetSpamDefault = new System.Windows.Forms.Button();
            this.btnEditProfile = new System.Windows.Forms.Button();
            this.btnAddProfile = new System.Windows.Forms.Button();
            this.btnDeleteProfile = new System.Windows.Forms.Button();
            this.lbProfiles = new System.Windows.Forms.ListBox();
            this.grpProfiles.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpProfiles
            // 
            this.grpProfiles.Controls.Add(this.btnSetHamDefault);
            this.grpProfiles.Controls.Add(this.btnSetSpamDefault);
            this.grpProfiles.Controls.Add(this.btnEditProfile);
            this.grpProfiles.Controls.Add(this.btnAddProfile);
            this.grpProfiles.Controls.Add(this.btnDeleteProfile);
            this.grpProfiles.Controls.Add(this.lbProfiles);
            this.grpProfiles.Location = new System.Drawing.Point(3, 0);
            this.grpProfiles.Name = "grpProfiles";
            this.grpProfiles.Size = new System.Drawing.Size(426, 232);
            this.grpProfiles.TabIndex = 2;
            this.grpProfiles.TabStop = false;
            this.grpProfiles.Text = "Profiles";
            // 
            // btnSetHamDefault
            // 
            this.btnSetHamDefault.Location = new System.Drawing.Point(330, 77);
            this.btnSetHamDefault.Name = "btnSetHamDefault";
            this.btnSetHamDefault.Size = new System.Drawing.Size(90, 23);
            this.btnSetHamDefault.TabIndex = 5;
            this.btnSetHamDefault.Text = "Ham Default";
            this.btnSetHamDefault.UseVisualStyleBackColor = true;
            this.btnSetHamDefault.Click += new System.EventHandler(this.btnSetHamDefault_Click);
            // 
            // btnSetSpamDefault
            // 
            this.btnSetSpamDefault.Location = new System.Drawing.Point(330, 48);
            this.btnSetSpamDefault.Name = "btnSetSpamDefault";
            this.btnSetSpamDefault.Size = new System.Drawing.Size(90, 23);
            this.btnSetSpamDefault.TabIndex = 4;
            this.btnSetSpamDefault.Text = "Spam Default";
            this.btnSetSpamDefault.UseVisualStyleBackColor = true;
            this.btnSetSpamDefault.Click += new System.EventHandler(this.btnSetDefault_Click);
            // 
            // btnEditProfile
            // 
            this.btnEditProfile.Location = new System.Drawing.Point(330, 135);
            this.btnEditProfile.Name = "btnEditProfile";
            this.btnEditProfile.Size = new System.Drawing.Size(90, 23);
            this.btnEditProfile.TabIndex = 3;
            this.btnEditProfile.Text = "Edit";
            this.btnEditProfile.UseVisualStyleBackColor = true;
            this.btnEditProfile.Click += new System.EventHandler(this.btnEditProfile_Click);
            // 
            // btnAddProfile
            // 
            this.btnAddProfile.Location = new System.Drawing.Point(330, 19);
            this.btnAddProfile.Name = "btnAddProfile";
            this.btnAddProfile.Size = new System.Drawing.Size(90, 23);
            this.btnAddProfile.TabIndex = 2;
            this.btnAddProfile.Text = "Add";
            this.btnAddProfile.UseVisualStyleBackColor = true;
            this.btnAddProfile.Click += new System.EventHandler(this.btnAddProfile_Click);
            // 
            // btnDeleteProfile
            // 
            this.btnDeleteProfile.Location = new System.Drawing.Point(330, 106);
            this.btnDeleteProfile.Name = "btnDeleteProfile";
            this.btnDeleteProfile.Size = new System.Drawing.Size(90, 23);
            this.btnDeleteProfile.TabIndex = 1;
            this.btnDeleteProfile.Text = "Delete";
            this.btnDeleteProfile.UseVisualStyleBackColor = true;
            this.btnDeleteProfile.Click += new System.EventHandler(this.btnDeleteProfile_Click);
            // 
            // lbProfiles
            // 
            this.lbProfiles.FormattingEnabled = true;
            this.lbProfiles.Location = new System.Drawing.Point(6, 19);
            this.lbProfiles.Name = "lbProfiles";
            this.lbProfiles.Size = new System.Drawing.Size(318, 199);
            this.lbProfiles.TabIndex = 0;
            this.lbProfiles.SelectedIndexChanged += new System.EventHandler(this.lbProfiles_SelectedIndexChanged);
            this.lbProfiles.DoubleClick += new System.EventHandler(this.lbProfiles_DoubleClick);
            this.lbProfiles.Leave += new System.EventHandler(this.lbProfiles_Leave);
            // 
            // ctlManageProfiles
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grpProfiles);
            this.Name = "ctlManageProfiles";
            this.Size = new System.Drawing.Size(432, 247);
            this.Load += new System.EventHandler(this.ctlManageProfiles_Load);
            this.grpProfiles.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpProfiles;
        private System.Windows.Forms.Button btnEditProfile;
        private System.Windows.Forms.Button btnAddProfile;
        private System.Windows.Forms.Button btnDeleteProfile;
        private System.Windows.Forms.ListBox lbProfiles;
        private System.Windows.Forms.Button btnSetSpamDefault;
        private System.Windows.Forms.Button btnSetHamDefault;
    }
}
