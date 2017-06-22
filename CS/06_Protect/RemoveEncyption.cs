using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveEncyption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //load the PPT with password
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Password.pptx", FileFormat.Pptx2010, "123456");

            //remove encryption
            presentation.RemoveEncryption();

            //save the document
            presentation.SaveToFile("RemoveEncryption.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RemoveEncryption.pptx");
        }
    }
}
