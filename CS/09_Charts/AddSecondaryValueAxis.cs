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

namespace AddSecondaryValueAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the chart from the PowerPoint file.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Add a secondary axis to display the value of Series 3.
            chart.Series[2].UseSecondAxis = true;

            //Set the grid line of secondary axis as invisible.
            chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;

            String result = "Result-AddSecondaryValueAxisToChart.pptx";

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