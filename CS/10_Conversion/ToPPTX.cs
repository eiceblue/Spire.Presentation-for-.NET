using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToPPTX
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToPPTX.ppt");

            //Save the PPT document to PPTX file format
            presentation.SaveToFile("ToPPTX.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("ToPPTX.pptx");
        }
    }
}