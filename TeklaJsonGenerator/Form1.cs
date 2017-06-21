using System;
using System.Diagnostics;
using System.Windows.Forms;
using Tekla.Structures.Dialog;

namespace TeklaJsonGenerator
{
    public partial class Form1 : ApplicationFormBase
    {
        public Form1()
        {
            InitializeComponent();

            if (!GetConnectionStatus())
            {
                MessageBox.Show("Tekla Structures is not running.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            JsonGenerator.Run();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
            Dispose();
        }

        private void linkLabelApp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = @"D:\Work\projects\_TFS\Common\MetalSpec\SpecGenerator\bin\Release\SpecGenerator.exe";
            myProcess.StartInfo.Arguments = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";
            myProcess.Start();
        }

        private void linkLabelJson_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "notepad.exe";
            myProcess.StartInfo.Arguments = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\TeklaMetalSpec.json";
            myProcess.Start();
        }
    }
}
