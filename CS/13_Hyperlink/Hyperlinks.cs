using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing.Animation;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace Hyperlinks
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

            //Set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);

            //Add new shape to PPT document
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 255, 120, 500, 280);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.Line.Width = 0;

            //Add some paragraphs with hyperlinks
            TextParagraph para1 = new TextParagraph();
            TextRange tr = new TextRange("E-iceblue");            
            tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            tr.Fill.SolidColor.Color = System.Drawing.Color.Blue;
            para1.TextRanges.Append(tr);
            para1.Alignment = TextAlignmentType.Center;
            shape.TextFrame.Paragraphs.Append(para1);
            shape.TextFrame.Paragraphs.Append(new TextParagraph());

            //Add some paragraphs with hyperlinks
            TextParagraph para2 = new TextParagraph();
            TextRange tr1 = new TextRange("Click to know more about Spire.Presentation.");
            tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
            para2.TextRanges.Append(tr1);
            shape.TextFrame.Paragraphs.Append(para2);
            shape.TextFrame.Paragraphs.Append(new TextParagraph());

            TextParagraph para3 = new TextParagraph();
            TextRange tr2 = new TextRange("Click to visit E-iceblue Home page.");
            tr2.ClickAction.Address = "https://www.e-iceblue.com/";
            para3.TextRanges.Append(tr2);
            shape.TextFrame.Paragraphs.Append(para3);
            shape.TextFrame.Paragraphs.Append(new TextParagraph());

            TextParagraph para4 = new TextParagraph();
            TextRange tr3 = new TextRange("Click to go to the forum to raise questions.");
            tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html";
            para4.TextRanges.Append(tr3);
            shape.TextFrame.Paragraphs.Append(para4);
            shape.TextFrame.Paragraphs.Append(new TextParagraph());

            TextParagraph para5 = new TextParagraph();
            TextRange tr4 = new TextRange("Click to contact our sales team via email.");
            tr4.ClickAction.Address = "mailto:sales@e-iceblue.com";
            para5.TextRanges.Append(tr4);
            shape.TextFrame.Paragraphs.Append(para5);
            shape.TextFrame.Paragraphs.Append(new TextParagraph());

            TextParagraph para6 = new TextParagraph();
            TextRange tr5 = new TextRange("Click to contact our support team via email.");
            tr5.ClickAction.Address = "mailto:support@e-iceblue.com";
            para6.TextRanges.Append(tr5);
            shape.TextFrame.Paragraphs.Append(para6);

            foreach (TextParagraph para in shape.TextFrame.Paragraphs)
            {
                if (!string.IsNullOrEmpty(para.Text))
                {
                    para.TextRanges[0].LatinFont = new TextFont("Lucida Sans Unicode");
                    para.TextRanges[0].FontHeight = 20;
                }

            }

            //Save the document
            presentation.SaveToFile("hyperlink_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("hyperlink_result.pptx");
        }
    }
}