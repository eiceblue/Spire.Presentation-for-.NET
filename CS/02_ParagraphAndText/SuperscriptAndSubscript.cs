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

namespace SuperscriptAndSubscript
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

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Az.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            //Add a shape 
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 100, 200, 50));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.TextFrame.Paragraphs.Clear();
            
            shape.AppendTextFrame("Test");
            TextRange tr = new TextRange("superscript");
            shape.TextFrame.Paragraphs[0].TextRanges.Append(tr);

            //Set superscript text
            shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = 30;
          
            TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.Black;
            textRange.FontHeight = 20;
            textRange.LatinFont = new TextFont("Lucida Sans Unicode");

            textRange = shape.TextFrame.Paragraphs[0].TextRanges[1];
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
            textRange.LatinFont = new TextFont("Lucida Sans Unicode");


            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 150, 200, 50));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.TextFrame.Paragraphs.Clear();

            shape.AppendTextFrame("Test");
            tr = new TextRange("subscript");
            shape.TextFrame.Paragraphs[0].TextRanges.Append(tr);

            //Set subscript text
            shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = -25;

            textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.Black;
            textRange.FontHeight = 20;
            textRange.LatinFont = new TextFont("Lucida Sans Unicode");

            textRange = shape.TextFrame.Paragraphs[0].TextRanges[1];
            textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
            textRange.LatinFont = new TextFont("Lucida Sans Unicode");


            string result = "SuperscriptAndSubscript_result.pptx";
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