using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Drawing.Animation;

namespace ApplyAnimationOnText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();

            //Get the first slide
            ISlide slide = ppt.Slides[0];

            //Set background image
            string ImageFile = @"..\..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Add a shape to the slide
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 150, 200, 100));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.LightBlue;
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.AppendTextFrame("This demo shows how to apply animation on text in PPT document.");

            //Apply animation to the text in shape
            AnimationEffect animation = shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Float);
            animation.SetStartEndParagraphs(0, 0);

            //Save the document
            string result = "ApplyAnimationOnText.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
		}
	
		private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}