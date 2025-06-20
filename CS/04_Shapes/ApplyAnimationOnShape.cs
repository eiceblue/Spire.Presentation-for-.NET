using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Drawing.Animation;

namespace ApplyAnimationOnShape
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

            //Set background Image
            string ImageFile = @"..\..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Insert a rectangle in the slide and fill the shape
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 150, 200, 80));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.LightBlue;
            shape.ShapeStyle.LineColor.Color = Color.White;
            shape.AppendTextFrame("Animated Shape");

            //Apply FadedSwivel animation effect to the shape
            shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel);

            //Save the document
            string result = "ApplyAnimationOnShape.pptx";
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