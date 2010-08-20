using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace SpamGrabberControl
{
    public partial class ctlAbout : UserControl
    {
        public ctlAbout()
        {
            InitializeComponent();
            this.linkLabel1.Links[0].LinkData = "SoftScan";
            this.linkLabel1.Links.Add(0, 8, "www.softscan.co.uk");

            this.linkLabel2.Links[0].LinkData = "SpamGrabber";
            this.linkLabel2.Links.Add(0, 11 ,"www.spamgrabber.org");
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void copyInfoButton_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(richTextBox1.Text);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string target = e.Link.LinkData as string;

            System.Diagnostics.Process.Start(target);
        }
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string target = e.Link.LinkData as string;

            System.Diagnostics.Process.Start(target);
        }

        private void lblVersion_Click(object sender, EventArgs e)
        {

        }

    }
}
