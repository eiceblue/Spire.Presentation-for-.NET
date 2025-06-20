using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace CreateWaterFallChart
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

            //Create a WaterFall chart to the first slide
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.WaterFall, new RectangleF(50, 50, 500, 400), false);

            //Set series text
            chart.ChartData[0, 1].Text = "Series 1";

            //Set category text
            string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7" };
            for (int i = 0; i < categories.Length; i++)
            {
                chart.ChartData[i + 1, 0].Text = categories[i];
            }

            //Fill data for chart
            double[] values = { 100, 20, 50, -40, 130, -60, 70 };
            for (int i = 0; i < values.Length; i++)
            {
                chart.ChartData[i + 1, 1].NumberValue = values[i];
            }

            //Set series labels
            chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];

            //Set categories labels 
            chart.Categories.CategoryLabels = chart.ChartData[1, 0, categories.Length, 0];

            //Assign data to series values
            chart.Series[0].Values = chart.ChartData[1, 1, values.Length, 1];

            //Operate the third datapoint of first series
            ChartDataPoint chartDataPoint = new ChartDataPoint(chart.Series[0]);
            chartDataPoint.Index = 2;
            chartDataPoint.SetAsTotal = true;
            chart.Series[0].DataPoints.Add(chartDataPoint);

            //Operate the sixth datapoint of first series
            ChartDataPoint chartDataPoint2 = new ChartDataPoint(chart.Series[0]);
            chartDataPoint2.Index = 5;
            chartDataPoint2.SetAsTotal = true;
            chart.Series[0].DataPoints.Add(chartDataPoint2);
            chart.Series[0].ShowConnectorLines = true;
            chart.Series[0].DataLabels.LabelValueVisible = true;

            chart.ChartLegend.Position = ChartLegendPositionType.Right;
            chart.ChartTitle.TextProperties.Text = "WaterFall";

            //Save the document
            string outputFile = "WaterFallChartResult.pptx";
            ppt.SaveToFile(outputFile, FileFormat.Pptx2013);
            ppt.Dispose();

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
