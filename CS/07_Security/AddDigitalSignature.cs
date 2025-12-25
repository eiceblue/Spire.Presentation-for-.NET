using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddDigitalSignature
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\AddDigitalSignature.pptx");

            //Add a digital signature,The parameters: string certificatePath, string certificatePassword, string comments, DateTime signTime
            ppt.AddDigitalSignature(@"..\..\..\..\..\..\Data\gary.pfx", "e-iceblue", "111", DateTime.Now);

            //Save the document
            ppt.SaveToFile("AddDigitalSignature_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddDigitalSignature_result.pptx");

            //Dispose
            ppt.Dispose();
        }
    }
}