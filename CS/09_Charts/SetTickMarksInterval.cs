using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Collections;
using System;
using System.IO;
using System.Windows.Forms;


namespace SetTickMarksInterval
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation ppt = new Presentation();
            string inputFile = @"..\..\..\..\..\..\Data\SetTickMarksInterval.pptx";
            ppt.LoadFromFile(inputFile);
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;
            IChartAxis chartAxis = chart.PrimaryCategoryAxis;
            chartAxis.TickMarkSpacing = 2;           
            //Save the document
            string outputFile = "SetTickMarksInterval_out.pptx";
            ppt.SaveToFile(outputFile, FileFormat.Pptx2013);

            //Launch the PPT file
            FileViewer(outputFile);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
