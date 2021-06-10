using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToPdfWithSpecificPageSize
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

            //Set A4 page size
            presentation.SlideSize.Type = SlideSizeType.A4;

            //Set landscape orientation
            presentation.SlideSize.Orientation = SlideOrienation.Landscape;

            String result = "ToPdfWithSpecifiedPageSize_result.pdf";
            //Save the PPT to PDF file format
            presentation.SaveToFile(result, FileFormat.PDF);

            Viewer(result);
        }

        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}