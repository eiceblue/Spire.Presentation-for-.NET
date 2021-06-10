using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace AutoVaryColorForPieChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT file
            Presentation ppt = new Presentation();

            RectangleF rect1 = new RectangleF(40, 100, 550, 320);

            //Add a pie chart
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Pie, rect1, false);
            chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            //Attach the data to chart
            string[] quarters = new string[] { "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr" };
            int[] sales = new int[] { 210, 320, 180, 500 };
            chart.ChartData[0, 0].Text = "Quarters";
            chart.ChartData[0, 1].Text = "Sales";
            for (int i = 0; i < quarters.Length; ++i)
            {
                chart.ChartData[i + 1, 0].Value = quarters[i];
                chart.ChartData[i + 1, 1].Value = sales[i];
            }

            chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
            chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
            chart.Series[0].Values = chart.ChartData["B2", "B5"];

         
            //Set whether auto vary color, default value is true
            chart.Series[0].IsVaryColor = false;

            chart.Series[0].Distance = 15;

            String result = "AutoVaryColorForPieChart_result.pptx";
            //Save the document
            ppt.SaveToFile(result, FileFormat.Pptx2010);

            PresentationDocViewer(result);
        }

        private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}