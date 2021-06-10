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
            //Create PPT document and load file.
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\FormatChartDataLabels.pptx");

            //Get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //Get the chart series
            ChartSeriesFormatCollection sers = chart.Series;
          
            //Initialize four instances of series label and set parameters of each label
            ChartDataLabel cd1 = sers[0].DataLabels.Add();                   
            cd1.PercentageVisible = true;
            cd1.TextFrame.Text = "Custom Datalabel1";
            cd1.TextFrame.TextRange.FontHeight = 12;
            cd1.TextFrame.TextRange.LatinFont =new TextFont("Lucida Sans Unicode");
            cd1.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            cd1.TextFrame.TextRange.Fill.SolidColor.Color= Color.Green;

            ChartDataLabel cd2 = sers[0].DataLabels.Add();
            cd2.Position = ChartDataLabelPosition.InsideEnd;
            cd2.PercentageVisible = true;
            cd2.TextFrame.Text = "Custom Datalabel2";
            cd2.TextFrame.TextRange.FontHeight = 10;
            cd2.TextFrame.TextRange.LatinFont = new TextFont("Arial");
            cd2.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            cd2.TextFrame.TextRange.Fill.SolidColor.Color = Color.OrangeRed;

            ChartDataLabel cd3 = sers[0].DataLabels.Add();
            cd3.Position = ChartDataLabelPosition.Center;
            cd3.PercentageVisible = true;
            cd3.TextFrame.Text = "Custom Datalabel3";
            cd3.TextFrame.TextRange.FontHeight = 14;
            cd3.TextFrame.TextRange.LatinFont = new TextFont("Calibri");
            cd3.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            cd3.TextFrame.TextRange.Fill.SolidColor.Color = Color.Blue;
            
            ChartDataLabel cd4 = sers[0].DataLabels.Add();
            cd4.Position = ChartDataLabelPosition.InsideBase;
            cd4.PercentageVisible = true;
            cd4.TextFrame.Text = "Custom Datalabel4";
            cd4.TextFrame.TextRange.FontHeight = 12;
            cd4.TextFrame.TextRange.LatinFont = new TextFont("Lucida Sans Unicode");
            cd4.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            cd4.TextFrame.TextRange.Fill.SolidColor.Color = Color.OliveDrab;

			//Save and launch the file 
            ppt.SaveToFile("FormatDataLable_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("FormatDataLable_result.pptx");
        }
    }
}
