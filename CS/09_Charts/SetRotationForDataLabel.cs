using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace SetRotationForDataLabel
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SetRotationForDataLabel.pptx");

            //Get chart on the first slide
            IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

            //Set the rotation angle for the datalabels of first serie
            for (int i = 0; i < Chart.Series[0].Values.Count; i++)
            {
                ChartDataLabel datalabel = Chart.Series[0].DataLabels.Add();
                datalabel.ID = i;
                datalabel.RotationAngle = 45;
            }

            String result = "SetRotationForDataLabel_out.pptx";

            //Save the document
            ppt.SaveToFile(result, FileFormat.Pptx2013);

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