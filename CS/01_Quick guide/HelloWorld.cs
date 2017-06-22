using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Spire.Presentation.Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //add new shape to PPT document
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 80, 500, 150);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

            //add text to shape
            shape.AppendTextFrame("Hello World!");

            //set the font and fill style of text
            TextRange textRange = shape.TextFrame.TextRange;
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.Black;
            textRange.FontHeight = 72;
            textRange.LatinFont = new TextFont("Myriad Pro Light");

            //save the document
            presentation.SaveToFile("hello.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("hello.pptx");
        }
    }
}