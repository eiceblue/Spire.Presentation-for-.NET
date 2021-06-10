using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FillShapeWithPattern
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

            //Add a rectangle
            RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 50, 100, 100, 100);
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect);

            //Set the pattern fill format 
            shape.Fill.FillType = FillFormatType.Pattern;
            shape.Fill.Pattern.PatternType = PatternFillType.Trellis;
            shape.Fill.Pattern.BackgroundColor.Color = Color.DarkGray;
            shape.Fill.Pattern.ForegroundColor.Color = Color.Yellow;

            //Set the fill format of line
            shape.Line.FillType = FillFormatType.Solid;
            shape.Line.SolidFillColor.Color = Color.Transparent;

            //Save the document
            string result = "FillShapeWithPattern_result.pptx";
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