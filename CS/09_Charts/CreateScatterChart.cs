using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateScatterChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void btnRun_Click(object sender, EventArgs e)
        {
            //Creat a presentation
            Presentation pres = new Presentation();


            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect2 = new RectangleF(0, 0, pres.SlideSize.Size.Width, pres.SlideSize.Size.Height);
            pres.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2);
            pres.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Insert a chart and set chart title and chart type
            RectangleF rect1 = new RectangleF(90, 100, 550, 320);
            IChart chart = pres.Slides[0].Shapes.AppendChart(ChartType.ScatterMarkers, rect1, false);
            chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            //Set chart data
            Double[] xdata = new Double[] { 2.7, 8.9, 10.0, 12.4 };
            Double[] ydata = new Double[] { 3.2, 15.3, 6.7, 8 };

            chart.ChartData[0, 0].Text = "X-Value";
            chart.ChartData[0, 1].Text = "Y-Value";

            for (Int32 i = 0; i < xdata.Length; ++i)
            {
                chart.ChartData[i + 1, 0].Value = xdata[i];
                chart.ChartData[i + 1, 1].Value = ydata[i];
            }

            //Set the series label
            chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];

            //Assign data to X axis, Y axis and Bubbles
            chart.Series[0].XValues = chart.ChartData["A2", "A5"];
            chart.Series[0].YValues = chart.ChartData["B2", "B5"];


            pres.SaveToFile("ScatterMarkerChart_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("ScatterMarkerChart_result.pptx");
        }
    }
}
