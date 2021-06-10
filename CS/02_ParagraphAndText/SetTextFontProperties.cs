using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;
using Spire.Presentation.Drawing;

namespace SetTextFontProperties
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
            shape.AppendTextFrame("Welcome to use Spire.Presentation");

            TextRange textRange = shape.TextFrame.TextRange;
            //Set the font
            textRange.LatinFont = new TextFont("Times New Roman");
            //Set bold property of the font
            textRange.IsBold = TriState.True;

            //Set italic property of the font
            textRange.IsItalic = TriState.True;

            //Set underline property of the font
            textRange.TextUnderlineType = TextUnderlineType.Single;

            //Set the height of the font
            textRange.FontHeight = 50;

            //Set the color of the font
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;

            string result = "SetTextFontProperties_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);
            Viewer(result);
        }

        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}