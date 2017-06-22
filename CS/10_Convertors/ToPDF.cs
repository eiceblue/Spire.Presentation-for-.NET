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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Presentation1.pptx");

            //save the PPT do PDF file format
            presentation.SaveToFile("ToPdf.pdf", FileFormat.PDF);
            System.Diagnostics.Process.Start("ToPdf.pdf");

        }
    }
}