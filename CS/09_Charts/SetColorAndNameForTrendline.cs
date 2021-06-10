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

namespace SetColorAndNameForTrendline
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a ppt document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SetColorAndNameForTrendline.pptx");

            //Find the first chart in the first Slide
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //Find the first trendline in the chart
            ITrendlines trendline = chart.Series[0].TrendLines[0] as ITrendlines;

            //Set name for trendline
            trendline.Name = "trendlineName";

            //Set color for trendline
            trendline.Line.FillType = FillFormatType.Solid;
            trendline.Line.SolidFillColor.Color = Color.Red;

            //Save the document
            ppt.SaveToFile("SetColorAndNameForTrendline_result.pptx", FileFormat.Pptx2010);

            //Launch the file
            OutputViewer("SetColorAndNameForTrendline_result.pptx");
        }

        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}