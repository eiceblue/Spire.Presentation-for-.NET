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

namespace AddCustomErrorBars
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample1.pptx");

            //Get the bubble chart on the first slide
            IChart bubbleChart = ppt.Slides[0].Shapes[0] as IChart;

            //Get X error bars of the first chart series
            IErrorBarsFormat errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat;
            //Specify error amount type as custom error bars
            errorBarsXFormat.ErrorBarvType = ErrorValueType.CustomErrorBars;
            //Set the minus and plus value of the X error bars
            errorBarsXFormat.MinusVal = 0.5f;
            errorBarsXFormat.PlusVal = 0.5f;

            //Get Y error bars of the first chart series
            IErrorBarsFormat errorBarsYFormat = bubbleChart.Series[0].ErrorBarsYFormat;
            //Specify error amount type as custom error bars
            errorBarsYFormat.ErrorBarvType = ErrorValueType.CustomErrorBars;
            //Set the minus and plus value of the Y error bars
            errorBarsYFormat.MinusVal = 1f;
            errorBarsYFormat.PlusVal = 1f;

            String result = "AddCustomErrorBars_result.pptx";
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