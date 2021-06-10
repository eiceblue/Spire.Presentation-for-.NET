using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace CopyChartWithinOnePPT
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

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the chart that is going to be copied.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Copy the chart from the first slide to the specified location of the second slide within the same document.
            ISlide slide1 = presentation.Slides.Append();
            slide1.Shapes.CreateChart(chart, new RectangleF(100, 100, 500, 300), 0);

            String result = "Result-CopyChartWithinAPptFile.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}