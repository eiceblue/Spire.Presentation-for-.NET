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

namespace MultipleParagraphs
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
            //Access the first slide
            ISlide slide = presentation.Slides[0];

            // Add an AutoShape of rectangle type
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 150, 500, 150);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);

            // Access TextFrame of the AutoShape
            ITextFrameProperties tf = shape.TextFrame;

            // Create Paragraphs and TextRanges with different text formats
            TextParagraph para0 = tf.Paragraphs[0];
            TextRange textRange1 = new TextRange();
            TextRange textRange2 = new TextRange();
            para0.TextRanges.Append(textRange1);
            para0.TextRanges.Append(textRange2);

            TextParagraph para1 = new TextParagraph();
            tf.Paragraphs.Append(para1);
            TextRange textRange11= new TextRange();
            TextRange textRange12 = new TextRange();
            TextRange textRange13 = new TextRange();
            para1.TextRanges.Append(textRange11);
            para1.TextRanges.Append(textRange12);
            para1.TextRanges.Append(textRange13);

            TextParagraph para2 = new TextParagraph();
            tf.Paragraphs.Append(para2);
            TextRange textRange21 = new TextRange();
            TextRange textRange22 = new TextRange();
            TextRange textRange23 = new TextRange();
            para2.TextRanges.Append(textRange21);
            para2.TextRanges.Append(textRange22);
            para2.TextRanges.Append(textRange23);

            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 3; j++)
                {
                    tf.Paragraphs[i].TextRanges[j].Text = "TextRange " + j.ToString();
                    if (j == 0)
                    {
                        tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid;
                        tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.LightBlue;
                        tf.Paragraphs[i].TextRanges[j].Format.IsBold = TriState.True;
                        tf.Paragraphs[i].TextRanges[j].FontHeight = 15;
                    }
                    else if (j == 1)
                    {
                        tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid;
                        tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.Blue;
                        tf.Paragraphs[i].TextRanges[j].Format.IsItalic = TriState.True;
                        tf.Paragraphs[i].TextRanges[j].FontHeight = 18;
                    }
                }


            string result = "MultipleParagraphs_result.pptx";
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