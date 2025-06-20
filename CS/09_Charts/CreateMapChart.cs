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

namespace CreatMapChart
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
            Presentation ppt = new Presentation();

            //Insert a Map chart to the first slide 
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Map, new RectangleF(50, 50, 450, 450), false);
            chart.ChartData[0, 1].Text = "series";
            
            //Define some data.
            string[] countries = { "China", "Russia", "France", "Mexico", "United States", "India", "Australia" };
            for (int i = 0; i < countries.Length; i++)
            {
                chart.ChartData[i + 1, 0].Text = countries[i];
            }
            int[] values = { 32, 20, 23, 17, 18, 6, 11 };
            for (int i = 0; i < values.Length; i++)
            {
                chart.ChartData[i + 1, 1].NumberValue = values[i];
            }
            chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];
            chart.Categories.CategoryLabels = chart.ChartData[1, 0, 7, 0];
            chart.Series[0].Values = chart.ChartData[1, 1, 7, 1];
            String result = "Result-CreateMapChart.pptx";

            //Save to file.
            ppt.SaveToFile(result, FileFormat.Pptx2013);

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