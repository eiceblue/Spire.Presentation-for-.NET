using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;
using System.IO;

namespace GetValuesAndUnitFromAxis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            //Create PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample2.pptx");

            //Get chart on the first slide
            IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

            //Get unit from primary category axis
            float MajorUnit = Chart.PrimaryCategoryAxis.MajorUnit;
            ChartBaseUnitType type = Chart.PrimaryCategoryAxis.MajorUnitScale;

            sb.Append(MajorUnit.ToString() + "\r\n");
            sb.Append(type.ToString() + "\r\n");


            //Get values from primary value axis
            float minValue = Chart.PrimaryValueAxis.MinValue;
            float maxValue = Chart.PrimaryValueAxis.MaxValue;

            sb.Append(minValue.ToString() + "\r\n");
            sb.Append(maxValue.ToString() + "\r\n");

                       
            String result = "GetValuesAndUnitFromAxis_result.txt";
            //Save the document
            File.WriteAllText(result, sb.ToString());

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