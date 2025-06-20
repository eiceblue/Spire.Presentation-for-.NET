using System;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetProjectionOfMapChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
 
            string inputFile = @"..\..\..\..\..\..\Data\SetProjectionOfMapChart.pptx";

            // Create Presentation object and load the file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(inputFile);

            // Get the chart
            IChart chart = ppt.Slides[0].Shapes[9] as IChart;

            // Get the type of projection
            ProjectionType type = chart.Series[0].ProjectionType;

            // Change the tpye of projection
            chart.Series[0].ProjectionType = ProjectionType.Robinson;

            // Save to file
            ppt.SaveToFile("SetProjectionOfMapChart_output2.pptx", FileFormat.Pptx2013);

            //Dispose
            ppt.Dispose();

            //System.Diagnostics.Process.Start("SetProjectionOfMapChart_output.pptx");

            this.Close();
        }

    }
}