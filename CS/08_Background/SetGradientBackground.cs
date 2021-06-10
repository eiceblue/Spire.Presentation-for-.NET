using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;
using Spire.Presentation.Drawing;

namespace SetGradientBackground
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

            //Load document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\PPTSample_N.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Set the background to gradient
            slide.SlideBackground.Type = BackgroundType.Custom;
            slide.SlideBackground.Fill.FillType = FillFormatType.Gradient;

            //Add gradient stops
            slide.SlideBackground.Fill.Gradient.GradientStops.Append(0.1f, Color.LightSeaGreen);
            slide.SlideBackground.Fill.Gradient.GradientStops.Append(0.7f, Color.LightCyan);

            //Set gradient shape type
            slide.SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear;

            //Set the angle
            slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45;

            //Save the document
            string result = "SetGradientBackground_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}