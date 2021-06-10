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

namespace SetDatapointColorInChart
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SetDatapointColorInChart.pptx");

            //Get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //Initialize an instances of dataPoint
            ChartDataPoint cdp1 = new ChartDataPoint(chart.Series[0]);
            
            //Specify the datapoint order
            cdp1.Index = 0;

            //Set the color of the datapoint
            cdp1.Fill.FillType = FillFormatType.Solid;
            cdp1.Fill.SolidColor.KnownColor = KnownColors.Orange;

            //Add the dataPoint to first series
            chart.Series[0].DataPoints.Add(cdp1);

            //Set the color for the other three data points
            ChartDataPoint cdp2 = new ChartDataPoint(chart.Series[0]);
            cdp2.Index = 1;
            cdp2.Fill.FillType = FillFormatType.Solid;
            cdp2.Fill.SolidColor.KnownColor = KnownColors.Gold;
            chart.Series[0].DataPoints.Add(cdp2);

            ChartDataPoint cdp3 = new ChartDataPoint(chart.Series[0]);
            cdp3.Index = 2;
            cdp3.Fill.FillType = FillFormatType.Solid;
            cdp3.Fill.SolidColor.KnownColor = KnownColors.MediumPurple;
            chart.Series[0].DataPoints.Add(cdp3);

            ChartDataPoint cdp4 = new ChartDataPoint(chart.Series[0]);
            cdp4.Index = 1;
            cdp4.Fill.FillType = FillFormatType.Solid;
            cdp4.Fill.SolidColor.KnownColor = KnownColors.ForestGreen;
            chart.Series[0].DataPoints.Add(cdp4);


            ppt.SaveToFile("SetDatapointColor_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("SetDatapointColor_result.pptx");
        }
    }
}
