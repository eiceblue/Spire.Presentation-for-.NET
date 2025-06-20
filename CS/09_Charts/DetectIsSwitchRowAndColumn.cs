using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Security.Policy;
using System.Text;
using System.Windows.Forms;

namespace DetectIsSwitchRowAndColumn
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample5.pptx");

            //Get the chart
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;

            //Detect whether the chart has "SwitchRowAndColumn" setting
            bool result = chart.IsSwitchRowAndColumn();

            MessageBox.Show("'SwitchRowAndColumn' value of the chart is " +result);
        }
    }
}
