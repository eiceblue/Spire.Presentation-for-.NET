using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Collections;
using Spire.Presentation.External.Pdf;

namespace ToPDFA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.pptx");

            //Save the PPT to PDF_A1A
            ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1A;
            string result = "ToPDF_A1A.pdf";
            ppt.SaveToFile(result, FileFormat.PDF);

            //Save the PPT to PDF_A1B
            ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
            result = "ToPDF_A1B.pdf";
            ppt.SaveToFile(result, FileFormat.PDF);

            //Save the PPT to PDF_A2A
            ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A2A;
            result = "ToPDF_A2A.pdf";
            ppt.SaveToFile(result, FileFormat.PDF);

            //View the document
            FileViewer(result);

        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}