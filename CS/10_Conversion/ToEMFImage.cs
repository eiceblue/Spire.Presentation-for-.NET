using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToEMFImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToEMFImage.pptx");

            //Save to EMF image
            presentation.Slides[0].SaveAsEMF("ToEMFImage.emf");
            System.Diagnostics.Process.Start("ToEMFImage.emf");
        }
    }
}
