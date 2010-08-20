namespace SpamGrabberControl
{
    partial class ctlProfileSettings
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
            this.submissionOptionsTab = new System.Windows.Forms.TabControl();
            this.tabReportAddress = new System.Windows.Forms.TabPage();
            this.submissionAddressBox = new System.Windows.Forms.GroupBox();
            this.btnFolderSelect = new System.Windows.Forms.Button();
            this.txtMoveToFolder = new System.Windows.Forms.TextBox();
            this.chkMoveToFolder = new System.Windows.Forms.CheckBox();
            this.chkMarkAsRead = new System.Windows.Forms.CheckBox();
            this.chkDeleteAfterReport = new System.Windows.Forms.CheckBox();
            this.afterReportLabel = new System.Windows.Forms.Label();
            this.btnDelBcc = new System.Windows.Forms.Button();
            this.btnAddBcc = new System.Windows.Forms.Button();
            this.btnDelTo = new System.Windows.Forms.Button();
            this.btnAddTo = new System.Windows.Forms.Button();
            this.bccLabel = new System.Windows.Forms.Label();
            this.toLabel = new System.Windows.Forms.Label();
            this.txtAddTo = new System.Windows.Forms.TextBox();
            this.txtAddBcc = new System.Windows.Forms.TextBox();
            this.lstBccAddresses = new System.Windows.Forms.ListBox();
            this.lstToAddresses = new System.Windows.Forms.ListBox();
            this.tabReportText = new System.Windows.Forms.TabPage();
            this.submissionTextBox = new System.Windows.Forms.GroupBox();
            this.messageBodyLabel = new System.Windows.Forms.Label();
            this.txtMessageBody = new System.Windows.Forms.RichTextBox();
            this.lblreportEndText = new System.Windows.Forms.Label();
            this.reportSubjectLabel = new System.Windows.Forms.Label();
            this.txtReportEndText = new System.Windows.Forms.TextBox();
            this.txtReportSubject = new System.Windows.Forms.TextBox();
            this.tabReportOptions = new System.Windows.Forms.TabPage();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.chkCleanHeaders = new System.Windows.Forms.CheckBox();
            this.chkSendReceive = new System.Windows.Forms.CheckBox();
            this.chkUseRfc = new System.Windows.Forms.CheckBox();
            this.chkSendMultiple = new System.Windows.Forms.CheckBox();
            this.chkKeepCopy = new System.Windows.Forms.CheckBox();
            this.chkAskVerify = new System.Windows.Forms.CheckBox();
            this.submissionOptionsTab.SuspendLayout();
            this.tabReportAddress.SuspendLayout();
            this.submissionAddressBox.SuspendLayout();
            this.tabReportText.SuspendLayout();
            this.submissionTextBox.SuspendLayout();
            this.tabReportOptions.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // submissionOptionsTab
            // 
            this.submissionOptionsTab.Controls.Add(this.tabReportAddress);
            this.submissionOptionsTab.Controls.Add(this.tabReportText);
            this.submissionOptionsTab.Controls.Add(this.tabReportOptions);
            this.submissionOptionsTab.Location = new System.Drawing.Point(3, 0);
            this.submissionOptionsTab.Name = "submissionOptionsTab";
            this.submissionOptionsTab.SelectedIndex = 0;
            this.submissionOptionsTab.Size = new System.Drawing.Size(426, 274);
            this.submissionOptionsTab.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.submissionOptionsTab.TabIndex = 0;
            // 
            // tabReportAddress
            // 
            this.tabReportAddress.Controls.Add(this.submissionAddressBox);
            this.tabReportAddress.Location = new System.Drawing.Point(4, 22);
            this.tabReportAddress.Name = "tabReportAddress";
            this.tabReportAddress.Padding = new System.Windows.Forms.Padding(3);
            this.tabReportAddress.Size = new System.Drawing.Size(418, 248);
            this.tabReportAddress.TabIndex = 0;
            this.tabReportAddress.Text = "Report Address";
            this.tabReportAddress.UseVisualStyleBackColor = true;
            // 
            // submissionAddressBox
            // 
            this.submissionAddressBox.Controls.Add(this.btnFolderSelect);
            this.submissionAddressBox.Controls.Add(this.txtMoveToFolder);
            this.submissionAddressBox.Controls.Add(this.chkMoveToFolder);
            this.submissionAddressBox.Controls.Add(this.chkMarkAsRead);
            this.submissionAddressBox.Controls.Add(this.chkDeleteAfterReport);
            this.submissionAddressBox.Controls.Add(this.afterReportLabel);
            this.submissionAddressBox.Controls.Add(this.btnDelBcc);
            this.submissionAddressBox.Controls.Add(this.btnAddBcc);
            this.submissionAddressBox.Controls.Add(this.btnDelTo);
            this.submissionAddressBox.Controls.Add(this.btnAddTo);
            this.submissionAddressBox.Controls.Add(this.bccLabel);
            this.submissionAddressBox.Controls.Add(this.toLabel);
            this.submissionAddressBox.Controls.Add(this.txtAddTo);
            this.submissionAddressBox.Controls.Add(this.txtAddBcc);
            this.submissionAddressBox.Controls.Add(this.lstBccAddresses);
            this.submissionAddressBox.Controls.Add(this.lstToAddresses);
            this.submissionAddressBox.Location = new System.Drawing.Point(6, 6);
            this.submissionAddressBox.Name = "submissionAddressBox";
            this.submissionAddressBox.Size = new System.Drawing.Size(406, 239);
            this.submissionAddressBox.TabIndex = 1;
            this.submissionAddressBox.TabStop = false;
            this.submissionAddressBox.Text = "Reporting address";
            // 
            // btnFolderSelect
            // 
            this.btnFolderSelect.Location = new System.Drawing.Point(257, 210);
            this.btnFolderSelect.Name = "btnFolderSelect";
            this.btnFolderSelect.Size = new System.Drawing.Size(71, 20);
            this.btnFolderSelect.TabIndex = 27;
            this.btnFolderSelect.Text = "Select";
            this.btnFolderSelect.UseVisualStyleBackColor = true;
            this.btnFolderSelect.Click += new System.EventHandler(this.btnFolderSelect_Click);
            // 
            // txtMoveToFolder
            // 
            this.txtMoveToFolder.Location = new System.Drawing.Point(118, 210);
            this.txtMoveToFolder.Name = "txtMoveToFolder";
            this.txtMoveToFolder.Size = new System.Drawing.Size(132, 20);
            this.txtMoveToFolder.TabIndex = 26;
            // 
            // chkMoveToFolder
            // 
            this.chkMoveToFolder.AutoSize = true;
            this.chkMoveToFolder.Location = new System.Drawing.Point(14, 210);
            this.chkMoveToFolder.Name = "chkMoveToFolder";
            this.chkMoveToFolder.Size = new System.Drawing.Size(97, 17);
            this.chkMoveToFolder.TabIndex = 25;
            this.chkMoveToFolder.Text = "Move to folder:";
            this.chkMoveToFolder.UseVisualStyleBackColor = true;
            this.chkMoveToFolder.CheckedChanged += new System.EventHandler(this.chkMoveToFolder_CheckedChanged);
            // 
            // chkMarkAsRead
            // 
            this.chkMarkAsRead.AutoSize = true;
            this.chkMarkAsRead.Location = new System.Drawing.Point(14, 193);
            this.chkMarkAsRead.Name = "chkMarkAsRead";
            this.chkMarkAsRead.Size = new System.Drawing.Size(88, 17);
            this.chkMarkAsRead.TabIndex = 24;
            this.chkMarkAsRead.Text = "Mark as read";
            this.chkMarkAsRead.UseVisualStyleBackColor = true;
            this.chkMarkAsRead.CheckedChanged += new System.EventHandler(this.chkMarkAsRead_CheckedChanged);
            // 
            // chkDeleteAfterReport
            // 
            this.chkDeleteAfterReport.AutoSize = true;
            this.chkDeleteAfterReport.Location = new System.Drawing.Point(14, 176);
            this.chkDeleteAfterReport.Name = "chkDeleteAfterReport";
            this.chkDeleteAfterReport.Size = new System.Drawing.Size(87, 17);
            this.chkDeleteAfterReport.TabIndex = 23;
            this.chkDeleteAfterReport.Text = "Delete e-mail";
            this.chkDeleteAfterReport.UseVisualStyleBackColor = true;
            this.chkDeleteAfterReport.CheckedChanged += new System.EventHandler(this.chkDeleteAfterReport_CheckedChanged);
            // 
            // afterReportLabel
            // 
            this.afterReportLabel.AutoSize = true;
            this.afterReportLabel.Location = new System.Drawing.Point(11, 159);
            this.afterReportLabel.Name = "afterReportLabel";
            this.afterReportLabel.Size = new System.Drawing.Size(123, 13);
            this.afterReportLabel.TabIndex = 22;
            this.afterReportLabel.Text = "After reporting messages";
            // 
            // btnDelBcc
            // 
            this.btnDelBcc.Location = new System.Drawing.Point(146, 120);
            this.btnDelBcc.Margin = new System.Windows.Forms.Padding(1, 0, 1, 5);
            this.btnDelBcc.Name = "btnDelBcc";
            this.btnDelBcc.Size = new System.Drawing.Size(27, 20);
            this.btnDelBcc.TabIndex = 19;
            this.btnDelBcc.Text = "<<";
            this.btnDelBcc.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnDelBcc.UseVisualStyleBackColor = true;
            this.btnDelBcc.Click += new System.EventHandler(this.btnDelBcc_Click);
            // 
            // btnAddBcc
            // 
            this.btnAddBcc.Location = new System.Drawing.Point(146, 96);
            this.btnAddBcc.Margin = new System.Windows.Forms.Padding(1, 0, 1, 5);
            this.btnAddBcc.Name = "btnAddBcc";
            this.btnAddBcc.Size = new System.Drawing.Size(27, 20);
            this.btnAddBcc.TabIndex = 18;
            this.btnAddBcc.Text = ">>";
            this.btnAddBcc.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnAddBcc.UseVisualStyleBackColor = true;
            this.btnAddBcc.Click += new System.EventHandler(this.btnAddBcc_Click);
            // 
            // btnDelTo
            // 
            this.btnDelTo.Location = new System.Drawing.Point(146, 47);
            this.btnDelTo.Margin = new System.Windows.Forms.Padding(1, 0, 1, 5);
            this.btnDelTo.Name = "btnDelTo";
            this.btnDelTo.Size = new System.Drawing.Size(27, 20);
            this.btnDelTo.TabIndex = 17;
            this.btnDelTo.Text = "<<";
            this.btnDelTo.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnDelTo.UseVisualStyleBackColor = true;
            this.btnDelTo.Click += new System.EventHandler(this.btnDelTo_Click);
            // 
            // btnAddTo
            // 
            this.btnAddTo.Location = new System.Drawing.Point(146, 23);
            this.btnAddTo.Margin = new System.Windows.Forms.Padding(1, 0, 1, 5);
            this.btnAddTo.Name = "btnAddTo";
            this.btnAddTo.Size = new System.Drawing.Size(27, 20);
            this.btnAddTo.TabIndex = 16;
            this.btnAddTo.Text = ">>";
            this.btnAddTo.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnAddTo.UseVisualStyleBackColor = true;
            this.btnAddTo.Click += new System.EventHandler(this.btnAddTo_Click);
            // 
            // bccLabel
            // 
            this.bccLabel.AutoSize = true;
            this.bccLabel.Location = new System.Drawing.Point(181, 77);
            this.bccLabel.Name = "bccLabel";
            this.bccLabel.Size = new System.Drawing.Size(28, 13);
            this.bccLabel.TabIndex = 10;
            this.bccLabel.Text = "BCC";
            // 
            // toLabel
            // 
            this.toLabel.AutoSize = true;
            this.toLabel.Location = new System.Drawing.Point(181, 8);
            this.toLabel.Name = "toLabel";
            this.toLabel.Size = new System.Drawing.Size(20, 13);
            this.toLabel.TabIndex = 9;
            this.toLabel.Text = "To";
            // 
            // txtAddTo
            // 
            this.txtAddTo.Location = new System.Drawing.Point(12, 24);
            this.txtAddTo.Name = "txtAddTo";
            this.txtAddTo.Size = new System.Drawing.Size(122, 20);
            this.txtAddTo.TabIndex = 7;
            // 
            // txtAddBcc
            // 
            this.txtAddBcc.Location = new System.Drawing.Point(12, 97);
            this.txtAddBcc.Name = "txtAddBcc";
            this.txtAddBcc.Size = new System.Drawing.Size(122, 20);
            this.txtAddBcc.TabIndex = 6;
            // 
            // lstBccAddresses
            // 
            this.lstBccAddresses.FormattingEnabled = true;
            this.lstBccAddresses.Location = new System.Drawing.Point(184, 97);
            this.lstBccAddresses.Name = "lstBccAddresses";
            this.lstBccAddresses.Size = new System.Drawing.Size(214, 43);
            this.lstBccAddresses.TabIndex = 5;
            // 
            // lstToAddresses
            // 
            this.lstToAddresses.FormattingEnabled = true;
            this.lstToAddresses.Location = new System.Drawing.Point(184, 25);
            this.lstToAddresses.Name = "lstToAddresses";
            this.lstToAddresses.Size = new System.Drawing.Size(214, 43);
            this.lstToAddresses.TabIndex = 4;
            // 
            // tabReportText
            // 
            this.tabReportText.AutoScroll = true;
            this.tabReportText.Controls.Add(this.submissionTextBox);
            this.tabReportText.Location = new System.Drawing.Point(4, 22);
            this.tabReportText.Name = "tabReportText";
            this.tabReportText.Padding = new System.Windows.Forms.Padding(3);
            this.tabReportText.Size = new System.Drawing.Size(418, 248);
            this.tabReportText.TabIndex = 1;
            this.tabReportText.Text = "Report Text";
            this.tabReportText.UseVisualStyleBackColor = true;
            // 
            // submissionTextBox
            // 
            this.submissionTextBox.Controls.Add(this.messageBodyLabel);
            this.submissionTextBox.Controls.Add(this.txtMessageBody);
            this.submissionTextBox.Controls.Add(this.lblreportEndText);
            this.submissionTextBox.Controls.Add(this.reportSubjectLabel);
            this.submissionTextBox.Controls.Add(this.txtReportEndText);
            this.submissionTextBox.Controls.Add(this.txtReportSubject);
            this.submissionTextBox.Location = new System.Drawing.Point(6, 6);
            this.submissionTextBox.Name = "submissionTextBox";
            this.submissionTextBox.Size = new System.Drawing.Size(406, 239);
            this.submissionTextBox.TabIndex = 1;
            this.submissionTextBox.TabStop = false;
            this.submissionTextBox.Text = "Submission text";
            // 
            // messageBodyLabel
            // 
            this.messageBodyLabel.AutoSize = true;
            this.messageBodyLabel.Location = new System.Drawing.Point(12, 89);
            this.messageBodyLabel.Name = "messageBodyLabel";
            this.messageBodyLabel.Size = new System.Drawing.Size(79, 13);
            this.messageBodyLabel.TabIndex = 5;
            this.messageBodyLabel.Text = "Message body:";
            // 
            // txtMessageBody
            // 
            this.txtMessageBody.Location = new System.Drawing.Point(12, 106);
            this.txtMessageBody.Name = "txtMessageBody";
            this.txtMessageBody.Size = new System.Drawing.Size(344, 90);
            this.txtMessageBody.TabIndex = 4;
            this.txtMessageBody.Text = "";
            this.txtMessageBody.TextChanged += new System.EventHandler(this.txtMessageBody_TextChanged);
            // 
            // lblreportEndText
            // 
            this.lblreportEndText.AutoSize = true;
            this.lblreportEndText.Location = new System.Drawing.Point(12, 56);
            this.lblreportEndText.Name = "lblreportEndText";
            this.lblreportEndText.Size = new System.Drawing.Size(79, 13);
            this.lblreportEndText.TabIndex = 3;
            this.lblreportEndText.Text = "End report text:";
            // 
            // reportSubjectLabel
            // 
            this.reportSubjectLabel.AutoSize = true;
            this.reportSubjectLabel.Location = new System.Drawing.Point(12, 24);
            this.reportSubjectLabel.Name = "reportSubjectLabel";
            this.reportSubjectLabel.Size = new System.Drawing.Size(79, 13);
            this.reportSubjectLabel.TabIndex = 2;
            this.reportSubjectLabel.Text = "Report subject:";
            // 
            // txtReportEndText
            // 
            this.txtReportEndText.Location = new System.Drawing.Point(146, 54);
            this.txtReportEndText.Name = "txtReportEndText";
            this.txtReportEndText.Size = new System.Drawing.Size(210, 20);
            this.txtReportEndText.TabIndex = 1;
            this.txtReportEndText.TextChanged += new System.EventHandler(this.txtReportEndText_TextChanged);
            // 
            // txtReportSubject
            // 
            this.txtReportSubject.Location = new System.Drawing.Point(146, 21);
            this.txtReportSubject.Name = "txtReportSubject";
            this.txtReportSubject.Size = new System.Drawing.Size(210, 20);
            this.txtReportSubject.TabIndex = 0;
            this.txtReportSubject.TextChanged += new System.EventHandler(this.txtReportSubject_TextChanged);
            // 
            // tabReportOptions
            // 
            this.tabReportOptions.Controls.Add(this.groupBox6);
            this.tabReportOptions.Location = new System.Drawing.Point(4, 22);
            this.tabReportOptions.Name = "tabReportOptions";
            this.tabReportOptions.Padding = new System.Windows.Forms.Padding(3);
            this.tabReportOptions.Size = new System.Drawing.Size(418, 248);
            this.tabReportOptions.TabIndex = 2;
            this.tabReportOptions.Text = "Report Options";
            this.tabReportOptions.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.chkCleanHeaders);
            this.groupBox6.Controls.Add(this.chkSendReceive);
            this.groupBox6.Controls.Add(this.chkUseRfc);
            this.groupBox6.Controls.Add(this.chkSendMultiple);
            this.groupBox6.Controls.Add(this.chkKeepCopy);
            this.groupBox6.Controls.Add(this.chkAskVerify);
            this.groupBox6.Location = new System.Drawing.Point(6, 6);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(406, 239);
            this.groupBox6.TabIndex = 2;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Report options";
            // 
            // chkCleanHeaders
            // 
            this.chkCleanHeaders.AutoSize = true;
            this.chkCleanHeaders.Location = new System.Drawing.Point(12, 95);
            this.chkCleanHeaders.Name = "chkCleanHeaders";
            this.chkCleanHeaders.Size = new System.Drawing.Size(188, 17);
            this.chkCleanHeaders.TabIndex = 9;
            this.chkCleanHeaders.Text = "Clean headers (not recommended)";
            this.chkCleanHeaders.UseVisualStyleBackColor = true;
            this.chkCleanHeaders.CheckedChanged += new System.EventHandler(this.chkCleanHeaders_CheckedChanged);
            // 
            // chkSendReceive
            // 
            this.chkSendReceive.AutoSize = true;
            this.chkSendReceive.Location = new System.Drawing.Point(12, 114);
            this.chkSendReceive.Name = "chkSendReceive";
            this.chkSendReceive.Size = new System.Drawing.Size(266, 17);
            this.chkSendReceive.TabIndex = 5;
            this.chkSendReceive.Text = "Attempt a send and receive after generating report ";
            this.chkSendReceive.UseVisualStyleBackColor = true;
            this.chkSendReceive.Visible = false;
            this.chkSendReceive.CheckedChanged += new System.EventHandler(this.chkSendReceive_CheckedChanged);
            // 
            // chkUseRfc
            // 
            this.chkUseRfc.AutoSize = true;
            this.chkUseRfc.Location = new System.Drawing.Point(12, 76);
            this.chkUseRfc.Name = "chkUseRfc";
            this.chkUseRfc.Size = new System.Drawing.Size(364, 17);
            this.chkUseRfc.TabIndex = 4;
            this.chkUseRfc.Text = "Use RFC822 for attachments (not compatable with Spamcop/SoftScan)";
            this.chkUseRfc.UseVisualStyleBackColor = true;
            this.chkUseRfc.CheckedChanged += new System.EventHandler(this.chkUseRfc_CheckedChanged);
            // 
            // chkSendMultiple
            // 
            this.chkSendMultiple.AutoSize = true;
            this.chkSendMultiple.Location = new System.Drawing.Point(12, 58);
            this.chkSendMultiple.Name = "chkSendMultiple";
            this.chkSendMultiple.Size = new System.Drawing.Size(189, 17);
            this.chkSendMultiple.TabIndex = 2;
            this.chkSendMultiple.Text = "Send multiple reports as one e-mail";
            this.chkSendMultiple.UseVisualStyleBackColor = true;
            this.chkSendMultiple.CheckedChanged += new System.EventHandler(this.chkSendMultiple_CheckedChanged);
            // 
            // chkKeepCopy
            // 
            this.chkKeepCopy.AutoSize = true;
            this.chkKeepCopy.Location = new System.Drawing.Point(12, 41);
            this.chkKeepCopy.Name = "chkKeepCopy";
            this.chkKeepCopy.Size = new System.Drawing.Size(256, 17);
            this.chkKeepCopy.TabIndex = 1;
            this.chkKeepCopy.Text = "Keep a copy of the report in my \'sent items\' folder";
            this.chkKeepCopy.UseVisualStyleBackColor = true;
            this.chkKeepCopy.CheckedChanged += new System.EventHandler(this.chkKeepCopy_CheckedChanged);
            // 
            // chkAskVerify
            // 
            this.chkAskVerify.AutoSize = true;
            this.chkAskVerify.Location = new System.Drawing.Point(12, 24);
            this.chkAskVerify.Name = "chkAskVerify";
            this.chkAskVerify.Size = new System.Drawing.Size(216, 17);
            this.chkAskVerify.TabIndex = 0;
            this.chkAskVerify.Text = "Ask for verification before sending report";
            this.chkAskVerify.UseVisualStyleBackColor = true;
            this.chkAskVerify.CheckedChanged += new System.EventHandler(this.chkAskVerify_CheckedChanged);
            // 
            // ctlProfileSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.submissionOptionsTab);
            this.Name = "ctlProfileSettings";
            this.Size = new System.Drawing.Size(432, 283);
            this.Load += new System.EventHandler(this.ctlProfileSettings_Load);
            this.submissionOptionsTab.ResumeLayout(false);
            this.tabReportAddress.ResumeLayout(false);
            this.submissionAddressBox.ResumeLayout(false);
            this.submissionAddressBox.PerformLayout();
            this.tabReportText.ResumeLayout(false);
            this.submissionTextBox.ResumeLayout(false);
            this.submissionTextBox.PerformLayout();
            this.tabReportOptions.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl submissionOptionsTab;
        private System.Windows.Forms.TabPage tabReportAddress;
        private System.Windows.Forms.GroupBox submissionAddressBox;
        private System.Windows.Forms.Button btnDelBcc;
        private System.Windows.Forms.Button btnAddBcc;
        private System.Windows.Forms.Button btnDelTo;
        private System.Windows.Forms.Button btnAddTo;
        private System.Windows.Forms.Label bccLabel;
        private System.Windows.Forms.Label toLabel;
        private System.Windows.Forms.TextBox txtAddTo;
        private System.Windows.Forms.TextBox txtAddBcc;
        private System.Windows.Forms.ListBox lstBccAddresses;
        private System.Windows.Forms.ListBox lstToAddresses;
        private System.Windows.Forms.TabPage tabReportText;
        private System.Windows.Forms.GroupBox submissionTextBox;
        private System.Windows.Forms.Label messageBodyLabel;
        private System.Windows.Forms.RichTextBox txtMessageBody;
        private System.Windows.Forms.Label lblreportEndText;
        private System.Windows.Forms.Label reportSubjectLabel;
        private System.Windows.Forms.TextBox txtReportEndText;
        private System.Windows.Forms.TextBox txtReportSubject;
        private System.Windows.Forms.TabPage tabReportOptions;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.CheckBox chkSendReceive;
        private System.Windows.Forms.CheckBox chkUseRfc;
        private System.Windows.Forms.CheckBox chkSendMultiple;
        private System.Windows.Forms.CheckBox chkKeepCopy;
        private System.Windows.Forms.CheckBox chkAskVerify;
        private System.Windows.Forms.Button btnFolderSelect;
        private System.Windows.Forms.TextBox txtMoveToFolder;
        private System.Windows.Forms.CheckBox chkMoveToFolder;
        private System.Windows.Forms.CheckBox chkMarkAsRead;
        private System.Windows.Forms.CheckBox chkDeleteAfterReport;
        private System.Windows.Forms.Label afterReportLabel;
        private System.Windows.Forms.CheckBox chkCleanHeaders;
    }
}
