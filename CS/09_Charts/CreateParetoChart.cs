using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Drawing;
using System.Windows.Forms;


namespace CreateParetoChart
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

            //Create a Pareto chart in first slide
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Pareto, new RectangleF(50, 50, 500, 400), false);

            //Set series text
            chart.ChartData[0, 1].Text = "Series 1";

            //Set category text
            string[] categories = { "Category 1", "Category 2", "Category 4", "Category 3", "Category 4", "Category 2", "Category 1",
                "Category 1", "Category 3", "Category 2", "Category 4", "Category 2", "Category 3",
                "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1",
                "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3",
                "Category 2", "Category 4", "Category 1"};
            for (int i = 0; i < categories.Length; i++)
            {
                chart.ChartData[i + 1, 0].Text = categories[i];
            }

            //Fill data for chart
            double[] values = { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            for (int i = 0; i < values.Length; i++)
            {
                chart.ChartData[i + 1, 1].NumberValue = values[i];
            }

            chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];
            chart.Categories.CategoryLabels = chart.ChartData[1, 0, categories.Length, 0];
            chart.Series[0].Values = chart.ChartData[1, 1, values.Length, 1];
            chart.PrimaryCategoryAxis.IsBinningByCategory = true;
            chart.Series[1].Line.FillFormat.FillType = FillFormatType.Solid;
            chart.Series[1].Line.FillFormat.SolidFillColor.Color = Color.Red;
            chart.ChartTitle.TextProperties.Text = "Pareto";
            chart.HasLegend = true;
            chart.ChartLegend.Position = ChartLegendPositionType.Bottom;

            //Save the document
            string outputFile = "ParetoChartResult.pptx";
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
