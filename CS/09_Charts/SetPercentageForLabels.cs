using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Collections;

namespace SetPercentageForLabels
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ColumnStacked.pptx");

            //Get the chart on the first slide
            IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

           float dataPontPercent = 0f;

           for (int i = 0; i < Chart.Series.Count; i++)
           {
               ChartSeriesDataFormat series = Chart.Series[i];
               //Get the total number
               float total = GetTotal(series.Values);
               for (int j = 0; j < series.Values.Count; j++)
               {
                 //Get the percent
                 dataPontPercent = float.Parse(series.Values[j].Text) / total * 100;
                 //Add datalabels
                 ChartDataLabel label = series.DataLabels.Add();
                 label.LabelValueVisible = true;
                 //Set the percent text for the label
                 label.TextFrame.Paragraphs[0].Text = String.Format("{0:F2} %", dataPontPercent);
                 label.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 12;
               }
           }

           String result = "SetPercentageForLabels_result.pptx";
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

        private float GetTotal(CellRanges ranges)
        {
            float total = 0;
            for (int i = 0; i < ranges.Count; i++)
            {
                total += float.Parse(ranges[i].Text);
            }

           return total;
        }
    }
}