using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetChartDataNumberFormat
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SetChartDataNumberFormat.pptx");

            //Get chart on the first slide
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //Set the number format for Axis
            chart.PrimaryValueAxis.NumberFormat = "#,##0.00";

            //Set the DataLabels format for Axis
            chart.Series[0].DataLabels.LabelValueVisible = true;
            chart.Series[0].DataLabels.PercentValueVisible = false;
            chart.Series[0].DataLabels.NumberFormat = "#,##0.00";
            chart.Series[0].DataLabels.HasDataSource = false;

            //Set the number format for ChartData
            for (int i = 1; i <= chart.Series[0].Values.Count; i++)
            {
                chart.ChartData[i, 1].NumberFormat = "#,##0.00";
            }

            String result = "SetChartDataNumberFormat_output.pptx";

            //Save the document
            ppt.SaveToFile(result, FileFormat.Pptx2013);

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