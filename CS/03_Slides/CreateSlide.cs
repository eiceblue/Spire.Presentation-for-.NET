using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace CreateSlide
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Add new slide
            presentation.Slides.Append();

            //Set the background image
            for (int i = 0; i < 2; i++)
            {
                string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
                RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
                presentation.Slides[i].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
                presentation.Slides[i].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;
            }

            //Add title
            RectangleF rec_title = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 70, 400, 50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.ShapeStyle.LineColor.Color = Color.White;
            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "E-iceblue";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //Append new shape
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 150, 600, 280));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.Line.FillType = FillFormatType.None;
            //Add text to shape
            shape.AppendTextFrame("Welcome to use Spire.Presentation for .NET.");

            //Add new paragraph
            TextParagraph pare = new TextParagraph();
            pare.Text = "";
            shape.TextFrame.Paragraphs.Append(pare);

            //Add new paragraph
            pare = new TextParagraph();
            pare.Text = "Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.";
            shape.TextFrame.Paragraphs.Append(pare);
          
            //Set the Font
            foreach (TextParagraph para in shape.TextFrame.Paragraphs)
            {
                para.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
                para.TextRanges[0].FontHeight = 24;
                para.TextRanges[0].Fill.FillType = FillFormatType.Solid;
                para.TextRanges[0].Fill.SolidColor.Color = Color.Black;
                para.Alignment = TextAlignmentType.Left;
            }

            //Append new shape - SixPointedStar
            shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.SixPointedStar, new RectangleF(100, 100, 100, 100));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.Orange;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Append new shape
            shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 250, 600, 50));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
         
            //Add text to shape
            shape.AppendTextFrame("This is newly added Slide.");
          
            //Set the Font
            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Myriad Pro");
            shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            shape.TextFrame.Paragraphs[0].Indent = 35;

            //Save the document
            presentation.SaveToFile("CreateSlide.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("CreateSlide.pptx");
        }
    }
}