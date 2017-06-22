using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Spire.Presentation.Demo
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

            //load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\source.pptx");

            //encrypy the document with password "test"
            presentation.Encrypt("test");

            //save the document
            presentation.SaveToFile("encrypt.pptx", FileFormat.Pptx2007);
            System.Diagnostics.Process.Start("encrypt.pptx");
        }
    }
}