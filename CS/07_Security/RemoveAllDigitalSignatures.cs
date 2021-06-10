using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace RemoveAllDigitalSignatures
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a ppt document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveAllDigitalSignatures.pptx");

            //Remove all digital signatures
            if (ppt.IsDigitallySigned == true)
            {
                ppt.RemoveAllDigitalSignatures();
            }
            //Save the document
            string output = @"RemoveAllDigitalSignatures_result.pptx";
            ppt.SaveToFile(output, FileFormat.Pptx2010);
            //Launch the file
            OutputViewer(output);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}