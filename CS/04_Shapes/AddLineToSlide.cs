using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddLineToSlide
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

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Add a line in the slide
            IAutoShape line=slide.Shapes.AppendShape(ShapeType.Line, new RectangleF(50, 100, 300, 0));

            //Set color of the line
            line.ShapeStyle.LineColor.Color = Color.Red;

            //Save the document
            string result = "AddLineToSlide_result.pptx";
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