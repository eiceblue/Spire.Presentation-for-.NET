using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace RemoveTickMarksOfAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the chart that need to be adjusted the number format and remove the tick marks.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Set percentage number format for the axis value of chart.
            chart.PrimaryValueAxis.NumberFormat = "0#\\%";

            //Remove the tick marks for value axis and category axis.
            chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkNone;
            chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkNone;
            chart.PrimaryCategoryAxis.MajorTickMark = TickMarkType.TickMarkNone;
            chart.PrimaryCategoryAxis.MinorTickMark = TickMarkType.TickMarkNone;

            String result = "Result-SetNumberFormatAndRemoveTickMarksOfChart.pptx";

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