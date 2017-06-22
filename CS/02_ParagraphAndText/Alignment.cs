using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;

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

            //set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //add title
            RectangleF rec_title = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 70, 400, 50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.ShapeStyle.LineColor.Color = Color.White;
            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "Alignment";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //append new shape
            RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width/2-250, 150, 500, 200);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.TextFrame.Paragraphs.RemoveAt(0);

            //add text to shape
            foreach (TextAlignmentType textAlign in Enum.GetValues(typeof(TextAlignmentType)))
            {
                //create a text range
                TextRange textRange = new TextRange("This text is " + textAlign.ToString());

                //create a new paragraph
                TextParagraph paragraph = new TextParagraph();

                //apend the text range
                paragraph.TextRanges.Append(textRange); 
                
                //set the alignment
                paragraph.Alignment = textAlign;

                //append to shape
                shape.TextFrame.Paragraphs.Append(paragraph);
            }


            //set the font and fill style
            foreach (TextParagraph paragraph in shape.TextFrame.Paragraphs)
            {
                paragraph.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
                paragraph.TextRanges[0].FontHeight = 24;
                paragraph.TextRanges[0].Fill.FillType = FillFormatType.Solid;
                paragraph.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            }

            //save the document
            presentation.SaveToFile("alignment.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("alignment.pptx");
        }
    }
}