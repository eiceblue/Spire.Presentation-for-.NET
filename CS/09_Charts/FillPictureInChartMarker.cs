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

namespace FillPictureInChartMarker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample4.pptx");

            //Get chart on the first slide
            IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

            //Load image file in ppt
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");
            IImageData IImage = ppt.Images.Append(image);

            //Create a ChartDataPoint object and specify the index
            ChartDataPoint dataPoint = new ChartDataPoint(Chart.Series[0]);
            dataPoint.Index = 0;

            //Fill picture in marker
            dataPoint.MarkerFill.Fill.FillType = FillFormatType.Picture;
            dataPoint.MarkerFill.Fill.PictureFill.Picture.EmbedImage = IImage;

            //Set marker size
            dataPoint.MarkerSize = 20;

            //Add the data point in series
            Chart.Series[0].DataPoints.Add(dataPoint);

            String result = "FillPictureInChartMarker_result.pptx";
            //Save the document
            ppt.SaveToFile(result, FileFormat.Pptx2010);

            //Launch the result file
            PPTDocViewer(result);
        }

        private void PPTDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }        
        }
    }
}