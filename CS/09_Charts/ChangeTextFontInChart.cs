using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ChangeTextFontInChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPTX file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeTextFontInChart.pptx");

            //Get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //Change the font of title
            chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Lucida Sans Unicode");
            chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue;
            chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 30;

            //Change the font of legend
            chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.DarkGreen;
            chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Lucida Sans Unicode");

            //Change the font of series
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10;
            chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Lucida Sans Unicode");
			
            ppt.SaveToFile("ChangeTextFontInChart_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("ChangeTextFontInChart_result.pptx");
        }
    }
}
