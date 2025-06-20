using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetCustomFonts
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

            //Add a new shape to the PPT document
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 80, 500, 150);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

            //Add text to the shape
            shape.AppendTextFrame("Hello World!");

            //Set the custom font folder
            presentation.SetCustomFontsFolder(@"E:\customFonts\");

            //Set the font and fill style of the text
            TextRange textRange = shape.TextFrame.TextRange;
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
            textRange.FontHeight = 66;
            textRange.LatinFont = new TextFont("Lucida Sans Unicode");

            //Save the document
            string result = @"CustomFonts_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2010);
            System.Diagnostics.Process.Start(result);
        }
    }
}