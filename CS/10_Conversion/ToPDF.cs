using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToPDF
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.pptx");

            //Save the PPT to PDF file format
            presentation.SaveToFile("ToPdf.pdf", FileFormat.PDF);
            System.Diagnostics.Process.Start("ToPdf.pdf");

        }
    }
}