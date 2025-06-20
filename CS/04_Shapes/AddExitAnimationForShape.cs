using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Presentation.Drawing.Animation;

namespace AddExitAnimationForShape
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
            IShape starShape = slide.Shapes.AppendShape(ShapeType.FivePointedStar, new RectangleF(250, 100, 200, 200));
            starShape.Fill.FillType = FillFormatType.Solid;
            starShape.Fill.SolidColor.KnownColor = KnownColors.LightBlue;

            //Add random bars effect to the shape
            AnimationEffect effect = slide.Timeline.MainSequence.AddEffect(starShape, AnimationEffectType.RandomBars);

            //Change effect type from entrance to exit
            effect.PresetClassType = TimeNodePresetClassType.Exit;

            //Save the document
            string result = "AddExitAnimationForShape.pptx";
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