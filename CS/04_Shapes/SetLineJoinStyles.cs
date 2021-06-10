using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetLineJoinStyles
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

            //Add three shapes
            IAutoShape shape1 = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 150, 150, 50));
            IAutoShape shape2 = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 150, 150, 50));
            IAutoShape shape3 = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(450, 150, 150, 50));

            //Fill shapes
            shape1.Fill.FillType = FillFormatType.Solid;
            shape1.Fill.SolidColor.Color = Color.CadetBlue;
            shape2.Fill.FillType = FillFormatType.Solid;
            shape2.Fill.SolidColor.Color = Color.CadetBlue;
            shape3.Fill.FillType = FillFormatType.Solid;
            shape3.Fill.SolidColor.Color = Color.CadetBlue;

            //Fill lines of shapes
            shape1.Line.FillType = FillFormatType.Solid;
            shape1.Line.SolidFillColor.Color = Color.DarkGray;
            shape2.Line.FillType = FillFormatType.Solid;
            shape2.Line.SolidFillColor.Color = Color.DarkGray;
            shape3.Line.FillType = FillFormatType.Solid;
            shape3.Line.SolidFillColor.Color = Color.DarkGray;

            //Set the line width
            shape1.Line.Width = 10;
            shape2.Line.Width = 10;
            shape3.Line.Width = 10;

            //Set the join styles of lines
            shape1.Line.JoinStyle = LineJoinType.Bevel;
            shape2.Line.JoinStyle = LineJoinType.Miter;
            shape3.Line.JoinStyle = LineJoinType.Round;

            //Add text in shapes
            shape1.TextFrame.Text = "Bevel Join Style";
            shape2.TextFrame.Text = "Miter Join Style";
            shape3.TextFrame.Text = "Round Join Style";

            //Save the document
            string result = "SetLineJoinStyles_result.pptx";
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