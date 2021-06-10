using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateBubbleChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT file.
            Presentation presentation = new Presentation();
			
            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect2 = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;
			
		    //Add bubble chart
            RectangleF rect1 = new RectangleF(90, 100, 550, 320);
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Bubble, rect1, false);
			
            //Chart title
            chart.ChartTitle.TextProperties.Text = "Bubble Chart";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            //Attach the data to chart
            Double[] xdata = new Double[] { 7.7, 8.9, 1.0, 2.4 };
            Double[] ydata = new Double[] { 15.2, 5.3, 6.7, 8 };
            Double[] size = new Double[] { 1.1, 2.4, 3.7, 4.8 };

            chart.ChartData[0, 0].Text = "X-Value";
            chart.ChartData[0, 1].Text = "Y-Value";
            chart.ChartData[0, 2].Text = "Size";

            for (Int32 i = 0; i < xdata.Length; ++i)
            {
                chart.ChartData[i + 1, 0].Value = xdata[i];
                chart.ChartData[i + 1, 1].Value = ydata[i];
                chart.ChartData[i + 1, 2].Value = size[i];
            }

            //Set series label
            chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];

            chart.Series[0].XValues = chart.ChartData["A2", "A5"];
            chart.Series[0].YValues = chart.ChartData["B2", "B5"];
            chart.Series[0].Bubbles.Add(chart.ChartData["C2"]);
            chart.Series[0].Bubbles.Add(chart.ChartData["C3"]);
            chart.Series[0].Bubbles.Add(chart.ChartData["C4"]);
            chart.Series[0].Bubbles.Add(chart.ChartData["C5"]);

            presentation.SaveToFile("BubbleChart_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("BubbleChart_result.pptx");
        }
    }
}
