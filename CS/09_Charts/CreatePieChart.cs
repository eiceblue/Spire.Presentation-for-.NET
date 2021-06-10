using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;

namespace CreatePieChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Insert a Pie chart to the first slide and set the chart title.
            RectangleF rect1 = new RectangleF(40, 100, 550, 320);
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Pie, rect1, false);
            chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            //Define some data.
            string[] quarters = new string[] { "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr" };
            int[] sales = new int[] { 210, 320, 180, 500 };

            //Append data to ChartData, which represents a data table where the chart data is stored.
            chart.ChartData[0, 0].Text = "Quarters";
            chart.ChartData[0, 1].Text = "Sales";
            for (int i = 0; i < quarters.Length; ++i)
            {
                chart.ChartData[i + 1, 0].Value = quarters[i];
                chart.ChartData[i + 1, 1].Value = sales[i];
            }

            //Set category labels, series label and series data.
            chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
            chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
            chart.Series[0].Values = chart.ChartData["B2", "B5"];

            //Add data points to series and fill each data point with different color.
            for (int i = 0; i < chart.Series[0].Values.Count; i++)
            {
                ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
                cdp.Index = i;
                chart.Series[0].DataPoints.Add(cdp);

            }
            chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.RosyBrown;
            chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.LightBlue;
            chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.LightPink;
            chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.MediumPurple;

            //Set the data labels to display label value and percentage value.
            chart.Series[0].DataLabels.LabelValueVisible = true;
            chart.Series[0].DataLabels.PercentValueVisible = true;

            String result = "Result-CreatePieChart.pptx";

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