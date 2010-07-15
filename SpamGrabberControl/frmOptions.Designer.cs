namespace SpamGrabberControl
{
    partial class frmOptions
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Global Settings");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Reporting Profiles");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("About");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOptions));
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.tvwOptions = new System.Windows.Forms.TreeView();
            this.about = new SpamGrabberControl.ctlAbout();
            this.submissionControl = new SpamGrabberControl.ctlProfileSettings();
            this.generalOptions = new SpamGrabberControl.ctlGeneralOptions();
            this.ManageProfilesControl = new SpamGrabberControl.ctlManageProfiles();
            this.SuspendLayout();
            // 
            // okButton
            // 
            this.okButton.Location = new System.Drawing.Point(409, 297);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 1;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(497, 297);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Location = new System.Drawing.Point(3, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(200, 100);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Report Options";
            // 
            // groupBox5
            // 
            this.groupBox5.Location = new System.Drawing.Point(3, 3);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(200, 100);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "General options";
            // 
            // groupBox7
            // 
            this.groupBox7.Location = new System.Drawing.Point(3, 3);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(200, 100);
            this.groupBox7.TabIndex = 0;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "General options";
            // 
            // tvwOptions
            // 
            this.tvwOptions.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.tvwOptions.Location = new System.Drawing.Point(12, 8);
            this.tvwOptions.Name = "tvwOptions";
            treeNode1.Checked = true;
            treeNode1.Name = "ndeGlobal";
            treeNode1.Text = "Global Settings";
            treeNode2.Name = "ndeProfiles";
            treeNode2.Text = "Reporting Profiles";
            treeNode3.Name = "ndeAbout";
            treeNode3.Text = "About";
            this.tvwOptions.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3});
            this.tvwOptions.Size = new System.Drawing.Size(122, 280);
            this.tvwOptions.TabIndex = 4;
            this.tvwOptions.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwOptions_AfterSelect);
            // 
            // about
            // 
            this.about.BackgroundImage = global::SpamGrabberControl.Resource.spamgrabber_7;
            this.about.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.about.Location = new System.Drawing.Point(146, 8);
            this.about.Name = "about";
            this.about.Size = new System.Drawing.Size(426, 274);
            this.about.TabIndex = 7;
            this.about.Visible = false;
            this.about.Load += new System.EventHandler(this.about_Load);
            // 
            // submissionControl
            // 
            this.submissionControl.Location = new System.Drawing.Point(146, 8);
            this.submissionControl.Name = "submissionControl";
            this.submissionControl.Profile = null;
            this.submissionControl.Size = new System.Drawing.Size(426, 283);
            this.submissionControl.TabIndex = 6;
            this.submissionControl.Visible = false;
            // 
            // generalOptions
            // 
            this.generalOptions.Location = new System.Drawing.Point(146, 8);
            this.generalOptions.Name = "generalOptions";
            this.generalOptions.Size = new System.Drawing.Size(426, 170);
            this.generalOptions.TabIndex = 5;
            // 
            // ManageProfilesControl
            // 
            this.ManageProfilesControl.Location = new System.Drawing.Point(146, 8);
            this.ManageProfilesControl.Name = "ManageProfilesControl";
            this.ManageProfilesControl.Size = new System.Drawing.Size(432, 248);
            this.ManageProfilesControl.TabIndex = 8;
            // 
            // frmOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(578, 324);
            this.Controls.Add(this.about);
            this.Controls.Add(this.tvwOptions);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.ManageProfilesControl);
            this.Controls.Add(this.submissionControl);
            this.Controls.Add(this.generalOptions);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmOptions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Options";
            this.Load += new System.EventHandler(this.frmOptions_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.TreeView tvwOptions;
        private SpamGrabberControl.ctlGeneralOptions generalOptions;
        private SpamGrabberControl.ctlProfileSettings submissionControl;
        private SpamGrabberControl.ctlAbout about;
        private ctlManageProfiles ManageProfilesControl;
    }
}