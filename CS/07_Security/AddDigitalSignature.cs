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

            //Load the certificate
            X509Certificate2 x509 = new X509Certificate2(@"..\..\..\..\..\..\Data\gary.pfx", "e-iceblue");

            //Add a digital signature
            ppt.AddDigitalSignature(x509, "111", DateTime.Now);

            //Save the document
            ppt.SaveToFile("AddDigitalSignature_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddDigitalSignature_result.pptx");
        }
    }
}