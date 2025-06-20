using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace CreateHistogramChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation ppt = new Presentation();

            //Add a Histogram chart
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Histogram, new RectangleF(50, 50, 500, 400), false);
            
            //Set series text
            chart.ChartData[0, 0].Text = "Series 1";

            //Fill data for chart
            double[] values = { 1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 8, 8, 8, 9, 9, 9, 12, 12, 13, 13, 17, 17, 17, 19, 19, 19, 25, 25, 25, 25, 25, 25, 25, 25, 29, 29, 29, 29, 32, 32, 33, 33, 35, 35, 41, 41, 44, 45, 49, 49 };
            for (int i = 0; i < values.Length; i++)
            {
                chart.ChartData[i + 1, 1].NumberValue = values[i];
            }

            //Set series label
            chart.Series.SeriesLabel = chart.ChartData[0, 0, 0, 0];

            //Set values for series
            chart.Series[0].Values = chart.ChartData[1, 0, values.Length, 0];

            chart.PrimaryCategoryAxis.NumberOfBins = 7;
            chart.PrimaryCategoryAxis.GapWidth = 20;
            //Chart title
            chart.ChartTitle.TextProperties.Text = "Histogram";
            chart.ChartLegend.Position = ChartLegendPositionType.Bottom;

            string outputFile = "histogramChartResult.pptx";
            //Save the document
            ppt.SaveToFile(outputFile, FileFormat.Pptx2013);

            //Launch the PPT file
            FileViewer(outputFile);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
