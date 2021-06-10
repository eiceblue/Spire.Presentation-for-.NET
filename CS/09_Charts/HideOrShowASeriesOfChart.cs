using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace HideOrShowASeriesOfChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the first slide.
            ISlide slide = presentation.Slides[0];

            //Get the first chart.
            IChart chart = slide.Shapes[0] as IChart;

            //Hide the first series of the chart.
            chart.Series[0].IsHidden = true;

            //Show the first series of the chart.
            //chart.Series[0].IsHidden = false;

            String result = "Result-HideOrShowASeriesOfChart.pptx";

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