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
using Spire.Presentation.Drawing;

namespace AddShadowEffectForDataLabel
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_3.pptx");

            //Get the chart.
            IChart chart = presentation.Slides[0].Shapes[0] as IChart;

            //Add a data label to the first chart series.
            ChartDataLabelCollection dataLabels = chart.Series[0].DataLabels;
            ChartDataLabel Label = dataLabels.Add();
            Label.LabelValueVisible = true;

            //Add outer shadow effect to the data label.
            Label.Effect.OuterShadowEffect = new OuterShadowEffect();

            //Set shadow color.
            Label.Effect.OuterShadowEffect.ColorFormat.Color = Color.Yellow;

            //Set blur.
            Label.Effect.OuterShadowEffect.BlurRadius = 5;

            //Set distance.
            Label.Effect.OuterShadowEffect.Distance = 10;

            //Set angle.
            Label.Effect.OuterShadowEffect.Direction = 90f;         

            String result = "Result-AddShadowEffectToChartDataLabels.pptx";

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