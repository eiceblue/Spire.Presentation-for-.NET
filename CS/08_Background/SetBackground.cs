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
            //Create PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SetBackground.pptx");

            //Set the background of the first slide to Gradient color
            presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom;
            presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Gradient;
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear;
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStyle = Spire.Presentation.Drawing.GradientStyle.FromTopLeftCorner;
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.Append(1f, KnownColors.SkyBlue);
            presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.Append(0f, KnownColors.White);

            //Set the background of the second slide to Solid color
            presentation.Slides[1].SlideBackground.Type = BackgroundType.Custom;
            presentation.Slides[1].SlideBackground.Fill.FillType = FillFormatType.Solid;
            presentation.Slides[1].SlideBackground.Fill.SolidColor.Color = Color.SkyBlue;

            presentation.Slides.Append();
            //Set the background of the third slide to picture
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[2].SlideBackground.Fill.FillType = FillFormatType.Picture;
            IEmbedImage image = presentation.Slides[2].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[2].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image as IImageData;


            //Save the document
            presentation.SaveToFile("SetBackground.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("SetBackground.pptx");
        }
    }
}
