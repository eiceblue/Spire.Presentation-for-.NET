using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetTextFontForChartTitle
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_3.pptx");

            //Get the chart.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Set the font for the text on chart title area.
            chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Arial Unicode MS");
            chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue;
            chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 50;

            String result = "Result-SetTextFontForChartTitle.pptx";

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