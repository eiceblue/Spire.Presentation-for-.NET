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
            para_title.Text = "Footer";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //add footer
            presentation.SetFooterText("Demo of Spire.Presentation");

            //set the footer visible
            presentation.SetFooterVisible(true);

            //set the page number visible
            presentation.SetSlideNoVisible(true);

            //set the date visible
            presentation.SetDateTimeVisible(true);

            //append new shape
            RectangleF rect2 = new RectangleF(presentation.SlideSize.Size.Width / 2 - 300, 155, 600, 250);
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect2);
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
           
            //add text to shape
            shape.AppendTextFrame("The sample demonstrates how to use Footer.");

            //append new paragraph
            shape.TextFrame.Paragraphs.Append(new TextParagraph());
            
            //add text to paragraph
            shape.TextFrame.Paragraphs[1].TextRanges.Append(new TextRange("Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine."));

            //set the Font
            foreach (TextParagraph para in shape.TextFrame.Paragraphs)
            {
                para.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
                para.TextRanges[0].FontHeight = 24;
                para.TextRanges[0].Fill.FillType = FillFormatType.Solid;
                para.TextRanges[0].Fill.SolidColor.Color = Color.Black;
                para.Alignment = TextAlignmentType.Left;
            }

            //save the document
            presentation.SaveToFile("HeaderAndFooter.pptx", FileFormat.Pptx2007);
            System.Diagnostics.Process.Start("HeaderAndFooter.pptx");
        }
    }
}