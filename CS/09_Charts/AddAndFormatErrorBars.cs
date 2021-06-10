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

namespace AddAndFormatErrorBars
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddAndFormatErrorBars.pptx");

            //Get the column chart on the first slide and set chart title.
            IChart columnChart = presentation.Slides[0].Shapes[0] as IChart;
           
            columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars"; 
            
           //Add Y (Vertical) Error Bars.

            //Get Y error bars of the first chart series.
            IErrorBarsFormat errorBarsYFormat1 = columnChart.Series[0].ErrorBarsYFormat;

            //Set end cap.
            errorBarsYFormat1.ErrorBarNoEndCap = false;

            //Specify direction.
            errorBarsYFormat1.ErrorBarSimType = ErrorBarSimpleType.Plus;

            //Specify error amount type.
            errorBarsYFormat1.ErrorBarvType = ErrorValueType.StandardError;

            //Set value.
            errorBarsYFormat1.ErrorBarVal = 0.3f;

            //Set line format.
            errorBarsYFormat1.Line.FillType = FillFormatType.Solid;
            errorBarsYFormat1.Line.SolidFillColor.Color = Color.MediumVioletRed;
            errorBarsYFormat1.Line.Width = 1;

            //Get the bubble chart on the second slide and set chart title.
            IChart bubbleChart = presentation.Slides[1].Shapes[0] as IChart;
            
            bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars";

            
             //Add X (Horizontal) and Y (Vertical) Error Bars.
            //Get X error bars of the first chart series.
            IErrorBarsFormat errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat;

            //Set end cap.
            errorBarsXFormat.ErrorBarNoEndCap = false;

            //Specify direction.
            errorBarsXFormat.ErrorBarSimType = ErrorBarSimpleType.Both;

            //Specify error amount type.
            errorBarsXFormat.ErrorBarvType = ErrorValueType.StandardError;

            //Set value.
            errorBarsXFormat.ErrorBarVal = 0.3f;

            //Get Y error bars of the first chart series.
            IErrorBarsFormat errorBarsYFormat2 = bubbleChart.Series[0].ErrorBarsYFormat;

            //Set end cap.
            errorBarsYFormat2.ErrorBarNoEndCap = false;

            //Specify direction.
            errorBarsYFormat2.ErrorBarSimType = ErrorBarSimpleType.Both;

            //Specify error amount type.
            errorBarsYFormat2.ErrorBarvType = ErrorValueType.StandardError;

            //Set value.
            errorBarsYFormat2.ErrorBarVal = 0.3f;

            String result = "Result-AddAndFormatErrorBars.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}