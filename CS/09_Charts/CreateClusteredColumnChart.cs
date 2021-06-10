using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Collections;

namespace CreateClusteredColumnChart
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
            Presentation presentation = new Presentation();

            //Add clustered column chart
            RectangleF rect1 = new RectangleF(90, 100, 550, 320);
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect1, false);

            //Chart title
            chart.ChartTitle.TextProperties.Text = "Clustered Column Chart";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            //Data for series
            Double[] Series1 = new Double[] { 7.7, 8.9, 1.0, 2.4 };
            Double[] Series2 = new Double[] { 15.2, 5.3, 6.7, 8 };
    
            //Set series text
            chart.ChartData[0, 1].Text = "Series1";
            chart.ChartData[0, 2].Text = "Series2";

            //Set category text
            chart.ChartData[1, 0].Text = "Category 1";
            chart.ChartData[2, 0].Text = "Category 2";
            chart.ChartData[3, 0].Text = "Category 3";
            chart.ChartData[4, 0].Text = "Category 4";
     
            //Fill data for chart
            for (Int32 i = 0; i < Series1.Length; ++i)
            {
                chart.ChartData[i + 1, 1].Value = Series1[i];
                chart.ChartData[i + 1, 2].Value = Series2[i];

            }

            //Set series label
            chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];
            //Set category label
            chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];

            //Set values for series
            chart.Series[0].Values = chart.ChartData["B2", "B5"];
            chart.Series[1].Values = chart.ChartData["C2", "C5"];

            String result = "CreateClusteredColumnChart_result.pptx";
            //Save the document
            presentation.SaveToFile(result, FileFormat.Pptx2010);

            //Launch the result file
            PPTDocViewer(result);
        }

        private void PPTDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }        
        }
    }
}