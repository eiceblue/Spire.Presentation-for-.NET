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

            //load the PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\source97.ppt");

            //save the PPT document to PPTX file format
            presentation.SaveToFile("ToPPTX.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("ToPPTX.pptx");
        }
    }
}