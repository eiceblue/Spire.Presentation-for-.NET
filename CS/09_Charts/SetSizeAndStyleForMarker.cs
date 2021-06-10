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

namespace SetSizeAndStyleForMarker
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SetSizeAndStyleForMarker.pptx");

            //Get the chart from the presentation.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            for (int i = 0; i < chart.Series[0].Values.Count; i++)
            {
                //Create a ChartDataPoint object and specify the index.
                ChartDataPoint dataPoint = new ChartDataPoint(chart.Series[0]);
                dataPoint.Index = i;

                //Set the fill color of the data marker.
                dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
                dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Yellow;

                //Set the line color of the data marker.
                dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
                dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.YellowGreen;

                //Set the size of the data marker.
                dataPoint.MarkerSize = 20;

                //Set the style of the data marker
                dataPoint.MarkerStyle = ChartMarkerType.Diamond;
                chart.Series[0].DataPoints.Add(dataPoint);
            }

            String result = "SetSizeAndStyleForMarker_out.pptx";

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