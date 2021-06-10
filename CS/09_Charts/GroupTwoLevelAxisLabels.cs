using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace GroupTwoLevelAxisLabels
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GroupTwoLevelAxisLabels.pptx");

            //Get the chart.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Get the category axis from the chart.
            IChartAxis chartAxis = chart.PrimaryCategoryAxis;

            //Group the axis labels that have the same first-level label.
            if (chartAxis.HasMultiLvlLbl)
            {
                chartAxis.IsMergeSameLabel = true;
            }    

            String result = "Result-GroupTwoLevelAxisLabels.pptx";

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