using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Drawing;
using System.Windows.Forms;


namespace CreateBoxAndWhiskerChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a PPT document
            Presentation ppt = new Presentation();

            // Insert a BoxAndWhisker chart to the first slide 
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.BoxAndWhisker, new RectangleF(50, 50, 500, 400), false);

            // Series labels
            string[] seriesLabel = { "Series 1", "Series 2", "Series 3" };
            for (int i = 0; i < seriesLabel.Length; i++)
            {
                chart.ChartData[0, i + 1].Text = "Series 1";
            }

            // Categories
            string[] categories = {"Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1",
                            "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2",
                            "Category 3", "Category 3", "Category 3", "Category 3", "Category 3"};
            for (int i = 0; i < categories.Length; i++)
            {
                chart.ChartData[i + 1, 0].Text = categories[i];
            }

            // Values
            double[,] values = new double[18, 3]{{-7,-3,-24},{-10,1,11},{-28,-6,34},{47,2,-21},{35,17,22},{-22,15,19},{17,-11,25},
                                        {-30,18,25},{49,22,56},{37,22,15},{-55,25,31},{14,18,22},{18,-22,36},{-45,25,-17},
                                        {-33,18,22},{18,2,-23},{-33,-22,10},{10,19,22}};
            for (int i = 0; i < seriesLabel.Length; i++)
            {
                for (int j = 0; j < categories.Length; j++)
                {
                    chart.ChartData[j + 1, i + 1].NumberValue = values[j, i];
                }
            }

            chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, seriesLabel.Length];
            chart.Categories.CategoryLabels = chart.ChartData[1, 0, categories.Length, 0];

            chart.Series[0].Values = chart.ChartData[1, 1, categories.Length, 1];
            chart.Series[1].Values = chart.ChartData[1, 2, categories.Length, 2];
            chart.Series[2].Values = chart.ChartData[1, 3, categories.Length, 3];

            chart.Series[0].ShowInnerPoints = false;
            chart.Series[0].ShowOutlierPoints = true;
            chart.Series[0].ShowMeanMarkers = true;
            chart.Series[0].ShowMeanLine = true;
            chart.Series[0].QuartileCalculationType = QuartileCalculation.ExclusiveMedian;

            chart.Series[1].ShowInnerPoints = false;
            chart.Series[1].ShowOutlierPoints = true;
            chart.Series[1].ShowMeanMarkers = true;
            chart.Series[1].ShowMeanLine = true;
            chart.Series[1].QuartileCalculationType = QuartileCalculation.InclusiveMedian;

            chart.Series[2].ShowInnerPoints = false;
            chart.Series[2].ShowOutlierPoints = true;
            chart.Series[2].ShowMeanMarkers = true;
            chart.Series[2].ShowMeanLine = true;
            chart.Series[2].QuartileCalculationType = QuartileCalculation.ExclusiveMedian;

            chart.HasLegend = true;
            chart.ChartTitle.TextProperties.Text = "BoxAndWhisker";
            chart.ChartLegend.Position = ChartLegendPositionType.Top;

            string outputFile = "result.pptx";
            //Save to file
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
