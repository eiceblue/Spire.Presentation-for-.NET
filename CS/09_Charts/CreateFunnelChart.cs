using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Drawing;
using System.Windows.Forms;


namespace CreateFunnelChart
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

            //Create a Funnel chart to the first slide
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Funnel, new RectangleF(50, 50, 550, 400), false);

            //Set series text
            chart.ChartData[0, 1].Text = "Series 1";

            //Set category text
            string[] categories = { "Website Visits", "Download", "Uploads", "Requested price", "Invoice sent", "Finalized" };
            for (int i = 0; i < categories.Length; i++)
            {
                chart.ChartData[i + 1, 0].Text = categories[i];
            }

            //Fill data for chart
            double[] values = { 50000, 47000, 30000, 15000, 9000, 5600 };
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

            //Set the chart title
            chart.ChartTitle.TextProperties.Text = "Funnel";

            string outputFile = "FunnelChartResult.pptx";
            //Save the document
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
