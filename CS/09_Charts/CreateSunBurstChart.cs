using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Drawing;
using System.Windows.Forms;


namespace CreateSunBurstChart
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

            //Create a SunBurst chart to the first slide
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.SunBurst, new RectangleF(50, 50, 500, 400), false);

            //Set series text
            chart.ChartData[0, 3].Text = "Series 1";

            //Set category text
            string[,] categories = {{"Branch 1","Stem 1","Leaf 1"},{"Branch 1","Stem 1","Leaf 2"},{"Branch 1","Stem 1", "Leaf 3"},
                 {"Branch 1","Stem 2","Leaf 4"},{"Branch 1","Stem 2","Leaf 5"},{"Branch 1","Leaf 6",null},{"Branch 1","Leaf 7", null},
                 {"Branch 2","Stem 3","Leaf 8"},{"Branch 2","Leaf 9",null},{"Branch 2","Stem 4","Leaf 10"},{"Branch 2","Stem 4","Leaf 11"},
                 {"Branch 2","Stem 5","Leaf 12"},{"Branch 3","Stem 5","Leaf 13"},{"Branch 3","Stem 6","Leaf 14"},{"Branch 3","Leaf 15",null}};
            for (int i = 0; i < 15; i++)
            {
                for (int j = 0; j < 3; j++)
                    chart.ChartData[i + 1, j].Value = categories[i, j];
            }

            //Fill data for chart
            double[] values = { 17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51 };
            for (int i = 0; i < values.Length; i++)
            {
                chart.ChartData[i + 1, 3].NumberValue = values[i];
            }

            //Set series labels
            chart.Series.SeriesLabel = chart.ChartData[0, 3, 0, 3];

            //Set categories labels 
            chart.Categories.CategoryLabels = chart.ChartData[1, 0, values.Length, 2];

            //Assign data to series values
            chart.Series[0].Values = chart.ChartData[1, 3, values.Length, 3];

            chart.Series[0].DataLabels.CategoryNameVisible = true;
            chart.ChartTitle.TextProperties.Text = "SunBurst";
            chart.HasLegend = true;
            chart.ChartLegend.Position = ChartLegendPositionType.Top;

            //Save the document
            string outputFile = "SunBurstChartResulr.pptx";
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
