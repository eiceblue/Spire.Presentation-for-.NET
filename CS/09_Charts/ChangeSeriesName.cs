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

namespace ChangeSeriesName
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample2.pptx");

            //Get chart on the first slide
            IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

            //Get the ranges of series label 
            CellRanges cr = Chart.Series.SeriesLabel;
            
            //Change the value
            cr[0].Value = "Changed series name";
            
            String result = "ChangeSeriesName_result.pptx";
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