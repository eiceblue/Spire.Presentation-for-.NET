using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SetSeriesLineColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SeriesLinesColor.pptx");

            //Get the first chart
            IShape shape = ppt.Slides[0].Shapes[0];
            if(shape is IChart)
            {
                IChart chart = (IChart)shape;
                TextLineFormat seriesLine = chart.SeriesLine;
                seriesLine.FillType = FillFormatType.Solid;

                //Set the color of seriesLine
                seriesLine.FillFormat.SolidFillColor.Color = Color.Red;
            }

            //Save the PPT document
            String result = "SeriesLinesColor_output.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
        }
        private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
