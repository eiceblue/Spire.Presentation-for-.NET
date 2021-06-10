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

namespace MultipleCategoryChart
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

            //Add line markers chart
            RectangleF rect1 = new RectangleF(90, 100, 550, 320);
            IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect1, false);

            //Chart title
            chart.ChartTitle.TextProperties.Text = "Muli-Category";
            chart.ChartTitle.TextProperties.IsCentered = true;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;


            //Data for series
            Double[] Series1 = new Double[] { 7.7, 8.9, 7, 6,7, 8 };
    
            //Set series text
            chart.ChartData[0, 2].Text = "Series1";

            //Set category text
            chart.ChartData[1, 0].Text = "Grp 1";
            chart.ChartData[3, 0].Text = "Grp 2";
            chart.ChartData[5, 0].Text = "Grp 3";           

            chart.ChartData[1, 1].Text = "A";
            chart.ChartData[2, 1].Text = "B";
            chart.ChartData[3, 1].Text = "C";
            chart.ChartData[4, 1].Text = "D";
            chart.ChartData[5, 1].Text = "E";
            chart.ChartData[6, 1].Text = "F";


            //Fill data for chart
            for (int i = 0; i < Series1.Length; ++i)
            {
                chart.ChartData[i + 1, 2].Value = Series1[i];

            }

            //Set series label
            chart.Series.SeriesLabel = chart.ChartData["C1", "C1"];
            //Set category label
            chart.Categories.CategoryLabels = chart.ChartData["A2", "B7"];
          
            //Set values for series
            chart.Series[0].Values = chart.ChartData["C2", "C7"];

            //Set if the category axis has multiple levels
            chart.PrimaryCategoryAxis.HasMultiLvlLbl = true;
            //Merge same label
            chart.PrimaryCategoryAxis.IsMergeSameLabel = true;

            String result = "MultipleCategoryChart_result.pptx";
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