using Spire.Presentation;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ShapeToSVG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Presentation object
            Presentation ppt = new Presentation();

            // Load a PowerPoint file ("ShapeToSVG.pptx")
            ppt.LoadFromFile(@"..\..\..\..\..\..\..\Data\toSVG.pptx");

            // Access the first slide in the presentation
            ISlide slide = ppt.Slides[0];

            // Initialize a counter for file naming
            int num = 0;

            // Iterate through each shape in the slide
            foreach (IShape shape in slide.Shapes)
            {
                // Save the shape as SVG format
                byte[] svgByte = shape.SaveAsSvg();

                // Create a new FileStream for writing the SVG content to a file
                FileStream fs = new FileStream("shape_" + num + ".svg", FileMode.Create);

                // Write the SVG content to the file
                fs.Write(svgByte, 0, svgByte.Length);

                // Close the FileStream
                fs.Close();

                // Increment the counter for the next file naming
                num++;
            }

            // Dispose of the Presentation object to release resources
            ppt.Dispose();
        }
    }
}
