using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetTextLinesInShape
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetLinesInShape.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

           StringBuilder sb = new StringBuilder();

            // Iterate the shapes in the slide
            for (int i=0;i<slide.Shapes.Count;i++)
            {
                // Get shape 
                IAutoShape shape = (IAutoShape)slide.Shapes[i];
                sb.Append("shape" + i + ":" + "\r\n");

                // Get text lines in the shape and get the text
                IList<LineText> lines = shape.TextFrame.GetLayoutLines();
                for (int j = 0; j < lines.Count; j++)
                {
                    sb.Append("line[" + j + "]:" + lines[j].Text + "\r\n");
                }
            }

            File.WriteAllText("GetLinesInShape.txt", sb.ToString());

            System.Diagnostics.Process.Start("GetLinesInShape.txt");

            presentation.Dispose();
        }
    }
}