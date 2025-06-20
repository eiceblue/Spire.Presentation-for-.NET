using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;
using Spire.Presentation.Charts;

namespace SetChartBorderstyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create Presentation
            Presentation presentation = new Presentation();

            //Load ppt file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample2.pptx");

            //Get chart on the first slide
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Set border style
            chart.Line.FillFormat.FillType = FillFormatType.Solid;
            chart.Line.FillFormat.SolidFillColor.Color = Color.Red;
            chart.BorderRoundedCorners = true;

            //Save the file
            String result = "SetChartBorderstyle_result.pptx";
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);

            //Launching the result file.
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}