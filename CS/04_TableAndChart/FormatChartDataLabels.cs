using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Collections;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FormatChartDataLabels
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document and load file.
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\PieChart.pptx");

            //get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //get the chart series
            ChartSeriesFormatCollection sers = chart.Series;

            //set the chart legend position to Right
            chart.ChartLegend.Position = ChartLegendPositionType.Right;
          
            //initialize four instances of series label and set parameters of each label
            ChartDataLabel cd1 = sers[0].DataLabels.Add();                   
            cd1.Position = ChartDataLabelPosition.Center;
            cd1.PercentageVisible = true;

            ChartDataLabel cd2 = sers[0].DataLabels.Add();
            cd2.PercentageVisible = true;
            cd2.Position = ChartDataLabelPosition.Center;

            ChartDataLabel cd3 = sers[0].DataLabels.Add();
            cd3.PercentageVisible = true;           
            cd3.Position = ChartDataLabelPosition.Center;

            ChartDataLabel cd4 = sers[0].DataLabels.Add();
            cd4.PercentageVisible = true;
            cd4.Position = ChartDataLabelPosition.Center;

            ppt.SaveToFile("FormatDataLable.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("FormatDataLable.pptx");
        }
    }
}
