using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateDoughnutChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a ppt document
            Presentation presentation = new Presentation();
            RectangleF rect = new RectangleF(80, 100, 550, 320);

			//Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect2 = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;
			
            //Add a Doughnut chart
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Doughnut, rect, false);
            chart.ChartTitle.TextProperties.Text = "Market share by country";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;

            string[] countries = new string[] { "Guba", "Mexico", "France", "German" };
            int[] sales = new int[] { 1800, 3000, 5100, 6200 };
            chart.ChartData[0, 0].Text = "Countries";
            chart.ChartData[0, 1].Text = "Sales";
            for (int i = 0; i < countries.Length; ++i)
            {
                chart.ChartData[i + 1, 0].Value = countries[i];
                chart.ChartData[i + 1, 1].Value = sales[i];
            }
            chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
            chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
            chart.Series[0].Values = chart.ChartData["B2", "B5"];

            for (int i = 0; i < chart.Series[0].Values.Count; i++)
            {
                ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
                cdp.Index = i;
                chart.Series[0].DataPoints.Add(cdp);
            }
            //Set the series color
            chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.LightBlue;
            chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.MediumPurple;
            chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.DarkGray;
            chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid;
            chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.DarkOrange;

            chart.Series[0].DataLabels.LabelValueVisible = true;
            chart.Series[0].DataLabels.PercentValueVisible = true;
            chart.Series[0].DoughnutHoleSize = 60;

            presentation.SaveToFile("DoughnutChart_result.pptx", FileFormat.Pptx2013);
            System.Diagnostics.Process.Start("DoughnutChart_result.pptx");

        }

        //Function to load data from XML file to DataTable
        private DataTable LoadData()
        {
            DataSet ds = new DataSet();
            ds.ReadXmlSchema(@"..\..\..\..\..\..\Data\data-schema.xml");
            ds.ReadXml(@"..\..\..\..\..\..\Data\data.xml");

            return ds.Tables[0];
        }

        //Function to load data from DataTable to IChart
        private void InitChartData(IChart chart, DataTable dataTable)
        {
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                chart.ChartData[0, c].Text = dataTable.Columns[c].Caption;
            }

            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                object[] data = dataTable.Rows[r].ItemArray;
                for (int c = 0; c < data.Length; c++)
                {
                    chart.ChartData[r + 1, c].Value = data[c];
                }
            }
        }
    }
}
