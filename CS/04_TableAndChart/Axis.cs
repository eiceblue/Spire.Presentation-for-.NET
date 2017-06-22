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

namespace UseSecondAxisAndFormatAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\BarChart.pptx");

            //get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //add a secondary axis to display the value of Series 3
            chart.Series[2].UseSecondAxis = true;

            //Set the grid line of secondary axis as invisible
            chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;

            //set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
            chart.PrimaryValueAxis.IsAutoMax = false;
            chart.PrimaryValueAxis.IsAutoMin = false;
            chart.SecondaryValueAxis.IsAutoMax = false;
            chart.SecondaryValueAxis.IsAutoMax = false;

            chart.PrimaryValueAxis.MinValue = 0f;
            chart.PrimaryValueAxis.MaxValue = 5.0f;
            chart.SecondaryValueAxis.MinValue = 0f;
            chart.SecondaryValueAxis.MaxValue = 1.0f;

            //set axis line format
            chart.PrimaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid;
            chart.SecondaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid;
            chart.PrimaryValueAxis.MinorGridLines.Width = 0.1f;
            chart.SecondaryValueAxis.MinorGridLines.Width = 0.1f;
            chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.LightGray;
            chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.LightGray;
            chart.PrimaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash;
            chart.SecondaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash;

            chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3f;
            chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.LightSkyBlue;
            chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3f;
            chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.LightSkyBlue;

            ppt.SaveToFile("secondAxis.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("secondAxis.pptx");
        }
    }
}
