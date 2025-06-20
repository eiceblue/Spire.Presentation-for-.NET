using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;
using Spire.Presentation.Charts;

namespace ChangesForTrendLineEquation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create Presentation
            Presentation presentation = new Presentation();

            //Load ppt file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\TrendlineEquation.pptx");

            //Get chart on the first slide
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Get the first trendline 
            ITrendlines trendline = chart.Series[0].TrendLines[0] as ITrendlines;

            //Change font size for trendline Equation text
            foreach (TextParagraph para in trendline.TrendLineLabel.TextFrameProperties.Paragraphs)
            {
                para.DefaultCharacterProperties.FontHeight = 20;
                foreach (Spire.Presentation.TextRange range in para.TextRanges)
                {
                    range.FontHeight = 20;
                }
            }

            //Change position for trendline Equation
            trendline.TrendLineLabel.OffsetX = -0.1f;
            trendline.TrendLineLabel.OffsetY = -0.05f;

            //Save the file
            String result = "ChangesForTrendLineEquation_result.pptx";
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);

            //Launching the result file.
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}