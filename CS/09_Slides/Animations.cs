using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing.Animation;
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
            para_title.Text = "Animation";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 36;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //set the animation of slide to Circle
            presentation.Slides[0].SlideShowTransition.Type = Spire.Presentation.Drawing.Transition.TransitionType.Circle;

            //append new shape - Triangle
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(100, 60, 100, 100));
           
            //set the color of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.Orange;
            shape.ShapeStyle.LineColor.Color = Color.White;
           
            //set the animation of shape
            shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Path4PointStar);

            //append new shape
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 200, 600, 260));
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
          
            //add text to shape
            shape.AppendTextFrame("This sample demonstrates how to set Animations in PPT using Spire.Presentation.");

            //add new paragraph
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
            presentation.SaveToFile("animations.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("animations.pptx");

        }
    }
}