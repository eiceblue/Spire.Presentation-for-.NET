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

namespace SetTextFontForLegendAndAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPonit document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the chart.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Set the font for the text on Chart Legend area.
            chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Green;
            chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Arial Unicode MS");

            //Set the font for the text on Chart Axis area.
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Arial Unicode MS");

            String result = "Result-SetTextFontOfChartLegendAndChartAxis.pptx";

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