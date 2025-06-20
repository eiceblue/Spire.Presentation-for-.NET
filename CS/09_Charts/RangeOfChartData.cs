using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace RangeOfChartData
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
            Presentation ppt = new Presentation();

            //Load PPT file 
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample2.pptx");

            //Create a StringBuilder object
            StringBuilder sb = new StringBuilder();

            //Get chart on the first slide
            IChart chart = ppt.Slides[0].Shapes[0] as IChart;
            if (chart != null)
            {
                int lastRow = chart.ChartData.LastRowIndex;
                int lastCol = chart.ChartData.LastColIndex;
                sb.AppendLine("lastRowIndex: " + lastRow + "\r\n" + "lastColIndex: " + lastCol);
            }

            //Save to txt file
            String result = "output.txt";
            File.WriteAllText(result, sb.ToString());

            System.Diagnostics.Process.Start(result);
        }
    }
}