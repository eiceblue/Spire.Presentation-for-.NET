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

namespace SetDatapointColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Chart.pptx");

            //get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //initialize an instances of dataPoint
            ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
            
            //specific the dataPoint
            cdp.Index = 2;

            //fill the dataPoint
            cdp.Fill.FillType = FillFormatType.Solid;
            cdp.Fill.SolidColor.KnownColor = KnownColors.Yellow;

            //add the dataPoint to first series
            chart.Series[0].DataPoints.Add(cdp);

            ppt.SaveToFile("SetDatapointColor.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("SetDatapointColor.pptx");
        }
    }
}
