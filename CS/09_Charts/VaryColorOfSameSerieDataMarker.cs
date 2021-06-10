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

namespace VaryColorOfSameSerieDataMarker
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\VaryColorsOfSameSeriesDataMarkers.pptx");

            //Get the chart from the presentation.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Create a ChartDataPoint object and specify the index.
            ChartDataPoint dataPoint = new ChartDataPoint(chart.Series[0]);
            dataPoint.Index = 0;

            //Set the fill color of the data marker.
            dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
            dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Red;

            //Set the line color of the data marker.
            dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
            dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Red;

            //Add the data point to the point collection of a series.
            chart.Series[0].DataPoints.Add(dataPoint);

            dataPoint = new ChartDataPoint(chart.Series[0]);
            dataPoint.Index = 1;
            dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
            dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Black;
            dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
            dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Black;
            chart.Series[0].DataPoints.Add(dataPoint);

            dataPoint = new ChartDataPoint(chart.Series[0]);
            dataPoint.Index = 2;
            dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
            dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Blue;
            dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
            dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Blue;
            chart.Series[0].DataPoints.Add(dataPoint);

            String result = "Result-VaryColorsOfSameSeriesDataMarkers.pptx";

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