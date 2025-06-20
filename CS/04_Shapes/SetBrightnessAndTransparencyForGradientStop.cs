using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetBrightnessAndTransparencyForGradientStop
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a Presentation object
            Presentation presentation = new Presentation();

            // Append new shape - BentUpArrow
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.BentUpArrow, new RectangleF(470, 300, 150, 100));

            // Set the color of shape
            shape.Fill.FillType = FillFormatType.Gradient;

            // Add gradient stops to create a gradient fill
            shape.Fill.Gradient.GradientStops.Append(0f, KnownColors.Olive);
            shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.PowderBlue);

            // Adjust the brightness and transparency of the first gradient stop
            shape.Fill.Gradient.GradientStops[0].Color.Brightness = 0.5f;
            shape.Fill.Gradient.GradientStops[0].Color.Transparency = 0.5f;

            // Set the line color of the shape
            shape.ShapeStyle.LineColor.Color = Color.White;

            // Specify the name for the output PowerPoint presentation file.
            string result = "SetBrightnessAndTransparencyForGradientStop.pptx";

            //Save the document
            presentation.SaveToFile(result, FileFormat.Pptx2010);

            //  Release resources
            presentation.Dispose();

            // Launch the saved file
            PresentationDocViewer(result);
        }
	
		private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}