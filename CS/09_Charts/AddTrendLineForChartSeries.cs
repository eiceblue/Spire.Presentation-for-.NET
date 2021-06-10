using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace AddTrendLineForChartSeries
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

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;
            ITrendlines it = chart.Series[0].AddTrendLine(TrendlinesType.Linear);

            //Set the trendline properties to determine what should be displayed.
            it.displayEquation = false;
            it.displayRSquaredValue = false;

            String result = "Result-AddTrendLineForChartSeries.pptx";

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