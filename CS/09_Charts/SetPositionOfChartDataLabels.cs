using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetPositionOfChartDataLabels
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

            //Get the chart.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Add data label to chart and set its id.
            ChartDataLabel label1 = chart.Series[0].DataLabels.Add();
            label1.ID = 0;

            //Set the default position of data label. This position is relative to the data markers.
            //label1.Position = ChartDataLabelPosition.OutsideEnd;

            //Set custom position of data label. This position is relative to the default position.
            label1.X = 0.1f;
            label1.Y = -0.1f;

            //Set label value visible
            label1.LabelValueVisible = true;

            //Set legend key invisible
            label1.LegendKeyVisible = false;

            //Set category name invisible
            label1.CategoryNameVisible = false;

            //Set series name invisible
            label1.SeriesNameVisible = false;

            //Set Percentage invisible
            label1.PercentageVisible = false;

            //Set border style and fill style of data label
            label1.Line.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            label1.Line.SolidFillColor.Color = Color.Blue;
            label1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            label1.Fill.SolidColor.Color = Color.Orange;

            String result = "Result-SetPositionOfChartDataLabels.pptx";

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