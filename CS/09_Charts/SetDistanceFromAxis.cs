using Spire.Presentation;
using Spire.Presentation.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace SetDistanceFromAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a ppt document
            Presentation ppt = new Presentation();

            //Append ColumnClustered chart
            IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, new RectangleF(50, 50, 400, 400));

            //Get the PrimaryCategory axis
            IChartAxis chartAxis = chart.PrimaryCategoryAxis;

            //Set "Distance from axis"
            chartAxis.LabelsDistance = 200;

            //Save the file
            ppt.SaveToFile("SetDistanceFromAxis.pptx", FileFormat.Pptx2013);

            //Launch and view the resulted PPTX file
            PresentationDocViewer("SetDistanceFromAxis.pptx");
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
