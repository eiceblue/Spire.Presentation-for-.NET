using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetChartDataLabelRange
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

            //Add a ColumnStacked chart
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnStacked, new RectangleF(100, 100, 500, 400));
            
            //Set data for the chart
            CellRange cellRange = chart.ChartData["F1"];
            cellRange.Text = "labelA";
            cellRange = chart.ChartData["F2"];
            cellRange.Text = "labelB";
            cellRange = chart.ChartData["F3"];
            cellRange.Text = "labelC";
            cellRange = chart.ChartData["F4"];
            cellRange.Text = "labelD";

            //Set data label ranges
            chart.Series[0].DataLabelRanges = chart.ChartData["F1", "F4"];

            //Add data label
            ChartDataLabel dataLabel1 = chart.Series[0].DataLabels.Add();
            dataLabel1.ID = 0;
            //Show the value
            dataLabel1.LabelValueVisible = false;
            //Show the label string
            dataLabel1.ShowDataLabelsRange = true;

            string result = "Result-SetChartDataLabelRange.pptx";
            //Save to file
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file
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