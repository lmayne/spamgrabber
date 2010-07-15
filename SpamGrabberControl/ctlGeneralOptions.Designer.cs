namespace SpamGrabberControl
{
    partial class ctlGeneralOptions
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
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.chkShowSelectButton = new System.Windows.Forms.CheckBox();
            this.chkReportToAllProfiles = new System.Windows.Forms.CheckBox();
            this.chkShowHamButton = new System.Windows.Forms.CheckBox();
            this.chkShowCopyButton = new System.Windows.Forms.CheckBox();
            this.chkEmbedStandardToolbar = new System.Windows.Forms.CheckBox();
            this.chkShowPreviewButtonBox = new System.Windows.Forms.CheckBox();
            this.chkShowSupportButton = new System.Windows.Forms.CheckBox();
            this.groupBox8.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.chkShowSupportButton);
            this.groupBox8.Controls.Add(this.chkShowSelectButton);
            this.groupBox8.Controls.Add(this.chkReportToAllProfiles);
            this.groupBox8.Controls.Add(this.chkShowHamButton);
            this.groupBox8.Controls.Add(this.chkShowCopyButton);
            this.groupBox8.Controls.Add(this.chkEmbedStandardToolbar);
            this.groupBox8.Controls.Add(this.chkShowPreviewButtonBox);
            this.groupBox8.Location = new System.Drawing.Point(3, 0);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(426, 160);
            this.groupBox8.TabIndex = 1;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "General options";
            // 
            // chkShowSelectButton
            // 
            this.chkShowSelectButton.AutoSize = true;
            this.chkShowSelectButton.Location = new System.Drawing.Point(12, 95);
            this.chkShowSelectButton.Name = "chkShowSelectButton";
            this.chkShowSelectButton.Size = new System.Drawing.Size(202, 17);
            this.chkShowSelectButton.TabIndex = 9;
            this.chkShowSelectButton.Text = "Show report to selected profile button";
            this.chkShowSelectButton.UseVisualStyleBackColor = true;
            this.chkShowSelectButton.CheckedChanged += new System.EventHandler(this.chkShowSelectButton_CheckedChanged);
            // 
            // chkReportToAllProfiles
            // 
            this.chkReportToAllProfiles.AutoSize = true;
            this.chkReportToAllProfiles.Location = new System.Drawing.Point(12, 114);
            this.chkReportToAllProfiles.Name = "chkReportToAllProfiles";
            this.chkReportToAllProfiles.Size = new System.Drawing.Size(241, 17);
            this.chkReportToAllProfiles.TabIndex = 8;
            this.chkReportToAllProfiles.Text = "Report spam and ham messages to all profiles";
            this.chkReportToAllProfiles.UseVisualStyleBackColor = true;
            this.chkReportToAllProfiles.CheckedChanged += new System.EventHandler(this.chkReportToAllProfiles_CheckedChanged);
            // 
            // chkShowHamButton
            // 
            this.chkShowHamButton.AutoSize = true;
            this.chkShowHamButton.Location = new System.Drawing.Point(12, 38);
            this.chkShowHamButton.Name = "chkShowHamButton";
            this.chkShowHamButton.Size = new System.Drawing.Size(186, 17);
            this.chkShowHamButton.TabIndex = 7;
            this.chkShowHamButton.Text = "Show report to default ham button";
            this.chkShowHamButton.UseVisualStyleBackColor = true;
            this.chkShowHamButton.CheckedChanged += new System.EventHandler(this.chkShowHam_CheckedChanged);
            // 
            // chkShowCopyButton
            // 
            this.chkShowCopyButton.AutoSize = true;
            this.chkShowCopyButton.Location = new System.Drawing.Point(12, 57);
            this.chkShowCopyButton.Name = "chkShowCopyButton";
            this.chkShowCopyButton.Size = new System.Drawing.Size(170, 17);
            this.chkShowCopyButton.TabIndex = 6;
            this.chkShowCopyButton.Text = "Show copy to clipboard button";
            this.chkShowCopyButton.UseVisualStyleBackColor = true;
            this.chkShowCopyButton.CheckedChanged += new System.EventHandler(this.chkShowCopyButton_CheckedChanged);
            // 
            // chkEmbedStandardToolbar
            // 
            this.chkEmbedStandardToolbar.AutoSize = true;
            this.chkEmbedStandardToolbar.Location = new System.Drawing.Point(12, 133);
            this.chkEmbedStandardToolbar.Name = "chkEmbedStandardToolbar";
            this.chkEmbedStandardToolbar.Size = new System.Drawing.Size(360, 17);
            this.chkEmbedStandardToolbar.TabIndex = 5;
            this.chkEmbedStandardToolbar.Text = "Embed program into standard commandbar (requires restarting Outlook)";
            this.chkEmbedStandardToolbar.UseVisualStyleBackColor = true;
            this.chkEmbedStandardToolbar.CheckedChanged += new System.EventHandler(this.chkEmbedStandardToolbar_CheckedChanged);
            // 
            // chkShowPreviewButtonBox
            // 
            this.chkShowPreviewButtonBox.AutoSize = true;
            this.chkShowPreviewButtonBox.Location = new System.Drawing.Point(12, 19);
            this.chkShowPreviewButtonBox.Name = "chkShowPreviewButtonBox";
            this.chkShowPreviewButtonBox.Size = new System.Drawing.Size(171, 17);
            this.chkShowPreviewButtonBox.TabIndex = 1;
            this.chkShowPreviewButtonBox.Text = "Show message preview button";
            this.chkShowPreviewButtonBox.UseVisualStyleBackColor = true;
            this.chkShowPreviewButtonBox.CheckedChanged += new System.EventHandler(this.chkShowPreviewButtonBox_CheckedChanged);
            // 
            // chkShowSupportButton
            // 
            this.chkShowSupportButton.AutoSize = true;
            this.chkShowSupportButton.Location = new System.Drawing.Point(12, 76);
            this.chkShowSupportButton.Name = "chkShowSupportButton";
            this.chkShowSupportButton.Size = new System.Drawing.Size(162, 17);
            this.chkShowSupportButton.TabIndex = 10;
            this.chkShowSupportButton.Text = "Show send to support button";
            this.chkShowSupportButton.UseVisualStyleBackColor = true;
            this.chkShowSupportButton.CheckedChanged += new System.EventHandler(this.chkShowSupportButton_CheckedChanged);
            // 
            // ctlGeneralOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox8);
            this.Name = "ctlGeneralOptions";
            this.Size = new System.Drawing.Size(432, 167);
            this.Load += new System.EventHandler(this.ctlGeneralOptions_Load);
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.CheckBox chkEmbedStandardToolbar;
        private System.Windows.Forms.CheckBox chkShowPreviewButtonBox;
        private System.Windows.Forms.CheckBox chkShowCopyButton;
        private System.Windows.Forms.CheckBox chkShowHamButton;
        private System.Windows.Forms.CheckBox chkReportToAllProfiles;
        private System.Windows.Forms.CheckBox chkShowSelectButton;
        private System.Windows.Forms.CheckBox chkShowSupportButton;
    }
}
