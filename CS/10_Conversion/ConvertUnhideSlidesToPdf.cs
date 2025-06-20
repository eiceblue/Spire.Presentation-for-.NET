using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ConvertUnhiddenSlidesToPdf
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\HideSlide1.pptx");

            //Convert the PPT unhidden slides to PDF file format 
            presentation.SaveToPdfOption.ContainHiddenSlides = false;
            string result = "ToPdf.pdf";
            presentation.SaveToFile(result, FileFormat.PDF);

            //View File
            DocumentViewer(result);
        }

        private static void DocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}