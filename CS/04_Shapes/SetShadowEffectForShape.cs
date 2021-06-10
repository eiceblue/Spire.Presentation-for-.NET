using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetShadowEffectForShape
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

            ISlide slide = ppt.Slides[0];

            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Add a shape to slide.
            RectangleF rect1 = new RectangleF(200, 150, 300, 120);
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect1);
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.LightBlue;
            shape.Line.FillType = FillFormatType.None;
            shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape.";
            shape.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.TextRange.Fill.SolidColor.Color = Color.Black;

            //Create an inner shadow effect through InnerShadowEffect object. 
            InnerShadowEffect innerShadow = new InnerShadowEffect();
            innerShadow.BlurRadius = 20;
            innerShadow.Direction = 0;
            innerShadow.Distance = 0;
            innerShadow.ColorFormat.Color = Color.Black;

            //Apply the shadow effect to shape.
            shape.EffectDag.InnerShadowEffect = innerShadow;

            //Save the document
            string result = "SetShadowEffectForShape.pptx";
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