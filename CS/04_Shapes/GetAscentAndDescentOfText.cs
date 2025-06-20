using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetAscentAndDescentOfText
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\..\Data\GetAscentAndDescentOfText.pptx");

            // Create a StringBuilder to store text information
            StringBuilder builder = new StringBuilder();

            // Access the first slide in the presentation
            ISlide slide = ppt.Slides[0];

            // Access the first AutoShape in the slide
            IAutoShape autoshape = slide.Shapes[0] as IAutoShape;

            // Retrieve the layout lines from the TextFrame of the AutoShape
            IList<LineText> lines = autoshape.TextFrame.GetLayoutLines();

            // Iterate through each layout line
            for (int i = 0; i < lines.Count; i++)
            {
                // Get the ascent and descent properties of the current line
                float ascent = lines[i].Ascent;
                float descent = lines[i].Descent;

                // Append information about the line, ascent, and descent to the StringBuilder
                builder.AppendLine("lines" + i + "\tascent: " + ascent + "\tdescent: " + descent);
            }

            // Specify the name for the result file
            string result = "GetAscentAndDescentOfText.txt";

            // Save to the text file
            File.WriteAllText(result, builder.ToString());

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