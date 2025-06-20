using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetTextPositionInShape
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

            // Load a PowerPoint file from a specified location
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\GetTextPositionInShape.pptx");

            // Create a StringBuilder to store text information
            StringBuilder sb = new StringBuilder();

            // Access the first slide in the presentation
            ISlide slide = ppt.Slides[0];

            // Iterate through all the shapes in the slide
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                // Get the current shape
                IShape shape = slide.Shapes[i];

                // Check if the shape is an AutoShape
                if (shape is IAutoShape)
                {
                    // Cast the shape to an AutoShape
                    IAutoShape autoshape = slide.Shapes[i] as IAutoShape;

                    // Get the text content of the AutoShape
                    string text = autoshape.TextFrame.Text;

                    // Obtain the text position information within the AutoShape
                    PointF point = autoshape.TextFrame.GetTextLocation();

                    // Append information about the shape, text, and location to the StringBuilder
                    sb.AppendLine("Shape " + i + "£º" + text + "\r\n" + "location£º" + point.ToString());
                }
            }
            // Specify the name for the result file
            string result = "GetTextPositionInShape.txt";

            // Append the collected information to a text file named "GetTextPositionInShape.txt"
            File.AppendAllText(result, sb.ToString());

            // Dispose of the Presentation object to release resources
            ppt.Dispose();
        
            // Launch the result file 
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