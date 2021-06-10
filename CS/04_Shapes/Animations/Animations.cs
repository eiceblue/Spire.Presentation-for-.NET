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

namespace Animations
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

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\Animations.pptx");

            //Add title
            RectangleF rec_title = new RectangleF(50, 200, 200, 50);
            IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
            shape_title.ShapeStyle.LineColor.Color = Color.Transparent;

            shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            TextParagraph para_title = new TextParagraph();
            para_title.Text = "Animations:";
            para_title.Alignment = TextAlignmentType.Center;
            para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
            para_title.TextRanges[0].FontHeight = 32;
            para_title.TextRanges[0].IsBold = TriState.True;
            para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(68, 68, 68);
            shape_title.TextFrame.Paragraphs.Append(para_title);

            //Set the animation of slide to Circle
            presentation.Slides[0].SlideShowTransition.Type = Spire.Presentation.Drawing.Transition.TransitionType.Circle;

            //Append new shape - Triangle
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(100, 280, 80, 80));

            //Set the color of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.CadetBlue;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Set the animation of shape
            shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Path4PointStar);

            //Append new shape - Rectangle and set animation
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(210, 280, 150, 80));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.CadetBlue;
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.AppendTextFrame("Animated Shape");
            shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel);

            //Append new shape - Cloud and set the animation
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Cloud, new RectangleF(390, 280, 80, 80));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.White;
            shape.ShapeStyle.LineColor.Color = Color.CadetBlue;
            shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedZoom);

            //Save the document
            presentation.SaveToFile("animations.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("animations.pptx");

        }
    }
}