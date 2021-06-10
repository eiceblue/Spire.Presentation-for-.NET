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

namespace Create100PercentStackedBarChart
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

            //Add a "Bar100PercentStacked" chart to the first slide.
            presentation.SlideSize.Type = SlideSizeType.Screen16x9;
            SizeF slidesize = presentation.SlideSize.Size;

            var slide = presentation.Slides[0];

            //Append a chart.
            RectangleF rect = new RectangleF(20, 20, slidesize.Width - 40, slidesize.Height - 40);
            IChart chart = slide.Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Bar100PercentStacked, rect);

            //Write data to the chart data.
            String[] columnlabels = { "Series 1", "Series 2", "Series 3" };

            //Insert the column labels.
            for (Int32 c = 0; c < columnlabels.Length; ++c)
                chart.ChartData[0, c + 1].Text = columnlabels[c];

            string[] rowlabels = { "Category 1", "Category 2", "Category 3" };

            //Insert the row labels.
            for (Int32 r = 0; r < rowlabels.Length; ++r)
                chart.ChartData[r + 1, 0].Text = rowlabels[r];

            double[,] values = new double[3, 3] { { 20.83233, 10.34323, -10.354667 }, { 10.23456, -12.23456, 23.34456 }, { 12.34345, -23.34343, -13.23232 } };

            //Insert the values.
            double value = 0.0;
            for (Int32 r = 0; r < rowlabels.Length; ++r)
            {
                for (Int32 c = 0; c < columnlabels.Length; ++c)
                {
                    value = Math.Round(values[r, c], 2);
                    chart.ChartData[r + 1, c + 1].Value = value;
                }
            }

            chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, columnlabels.Length];
            chart.Categories.CategoryLabels = chart.ChartData[1, 0, rowlabels.Length, 0];

            //Set the position of category axis.
            chart.PrimaryCategoryAxis.Position = AxisPositionType.Left;
            chart.SecondaryCategoryAxis.Position = AxisPositionType.Left;
            chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow;

            //Set the data, font and format for the series of each column.
            for (Int32 c = 0; c < columnlabels.Length; ++c)
            {
                chart.Series[c].Values = chart.ChartData[1, c + 1, rowlabels.Length, c + 1];
                chart.Series[c].Fill.FillType = FillFormatType.Solid;
                chart.Series[c].InvertIfNegative = false;

                for (Int32 r = 0; r < rowlabels.Length; ++r)
                {
                    var label = chart.Series[c].DataLabels.Add();
                    label.LabelValueVisible = true;
                    chart.Series[c].DataLabels[r].HasDataSource = false;
                    chart.Series[c].DataLabels[r].NumberFormat = "0#\\%";
                    chart.Series[c].DataLabels.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 12;
                }
            }

            //Set the color of the Series.
            chart.Series[0].Fill.SolidColor.Color = Color.YellowGreen;
            chart.Series[1].Fill.SolidColor.Color = Color.Red;
            chart.Series[2].Fill.SolidColor.Color = Color.Green;

            TextFont font = new TextFont("Tw Cen MT");

            //Set the font and size for chartlegend.
            for (int k = 0; k < chart.ChartLegend.EntryTextProperties.Length; k++)
            {
                chart.ChartLegend.EntryTextProperties[k].LatinFont = font;
                chart.ChartLegend.EntryTextProperties[k].FontHeight = 20;
            }

            String result = "Result-Create100PercentStackedBarChart.pptx";

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