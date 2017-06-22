using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetBackground
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

            //add new slide
            presentation.Slides.Append();
            //add new slide
            presentation.Slides.Append();

            //set the background of the first slide to Gradient color
            presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom;
            presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Gradient;
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear;
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStyle = Spire.Presentation.Drawing.GradientStyle.FromCorner1;
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.Append(1f, KnownColors.LightGreen);
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.Append(0f, KnownColors.White);

            //set the background of the second slide to Solid color
            presentation.Slides[1].SlideBackground.Type = BackgroundType.Custom;
            presentation.Slides[1].SlideBackground.Fill.FillType = FillFormatType.Solid;
            presentation.Slides[1].SlideBackground.Fill.SolidColor.Color = Color.DarkSeaGreen;

            //set the background of the third slide to picture
            string ImageFile = @"..\..\..\..\..\..\Data\background.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[2].SlideBackground.Fill.FillType = FillFormatType.Picture;
            IEmbedImage image = presentation.Slides[2].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[2].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image as IImageData;

            //add shape and fill it with text in slides
            IAutoShape shape;
            TextRange textRange;
            for (int i = 0; i < 3; i++)
            {
                shape = presentation.Slides[i].Shapes.AppendShape(ShapeType.Rectangle,
                new RectangleF(50, 70, 600, 100));
                shape.ShapeStyle.LineColor.Color = Color.White;
                shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

                shape.AppendTextFrame("Demonstrates to how to set the background style of slides.");

                //set the Font and fill style
                textRange = shape.TextFrame.TextRange;
                textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                textRange.Fill.SolidColor.Color = Color.Black;
                textRange.LatinFont = new TextFont("Arial Black");
            }
            //save the document
            presentation.SaveToFile("background.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("background.pptx");
        }
    }
}
