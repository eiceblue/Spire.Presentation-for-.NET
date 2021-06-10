using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ConvertODPtoPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {

            Presentation presentation = new Presentation();

            //Load ODP file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\toPdf.odp",FileFormat.ODP);

            String result = "ConvertODPtoPDF_result.pdf";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.PDF);

            //Launch the PowerPoint file.
            DocumentViewer(result);
        }

        private void DocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}