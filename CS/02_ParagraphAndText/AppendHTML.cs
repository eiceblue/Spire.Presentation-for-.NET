using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AppendHTML
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
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\AppendHTML.pptx");
            //Add a shape 
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 100, 200, 200));

            //Clear default paragraphs 
            shape.TextFrame.Paragraphs.Clear();

            string code = "<html><body><p>This is a paragraph</p></body></html>";

            //Append HTML, and generate a paragraph with default style in PPT document.
            shape.TextFrame.Paragraphs.AddFromHtml(code);
            string codeColor = "<html><body><p style=\" color:black \">This is a paragraph</p></body></html>";
            //Append HTML with black setting
            shape.TextFrame.Paragraphs.AddFromHtml(codeColor);

            //Add another shape
            IAutoShape shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(350, 100, 200, 200));

            //Clear default paragraph 
            shape1.TextFrame.Paragraphs.Clear();

            //Change the fill format of shape
            shape1.Fill.FillType = FillFormatType.Solid;
            shape1.Fill.SolidColor.Color = Color.White;

            //Append HTML
            shape1.TextFrame.Paragraphs.AddFromHtml(code);
            TextParagraph par = shape1.TextFrame.Paragraphs[0];
            //Change the fill color for paragraph
            foreach (TextRange tr in par.TextRanges)
            {
                tr.Fill.FillType = FillFormatType.Solid;
                tr.Fill.SolidColor.Color = Color.Black;
            }

            ppt.SaveToFile("AppendHTML_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AppendHTML_result.pptx");

        }
    }
}
