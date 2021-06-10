using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetTickMarkLabelsOnCategoryAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPonit document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_3.pptx");

            //Get the chart from the PowerPoint slide.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Rotate tick labels.
            chart.PrimaryCategoryAxis.TextRotationAngle = 45;

            //Specify interval between labels.
            chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = false;
            chart.PrimaryCategoryAxis.TickLabelSpacing = 2;

            //Change position.
            chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh;
            
            String result = "Result-SetTickMarkLabelsOnCategoryAxis.pptx";

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