using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetShadowEffectForText
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

            //Set background image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            ppt.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Get reference of the slide
            ISlide slide = ppt.Slides[0];

            //Add a new rectangle shape to the first slide
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(120, 100, 450, 200));
            shape.Fill.FillType = FillFormatType.None;

            //Add the text to the shape and set the font for the text
            shape.AppendTextFrame("Text shading on slides");
            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Black");
            shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 21;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;

            ////Add inner shadow and set all necessary parameters
            //InnerShadowEffect Shadow = InnerShadowEffect();

            //Add outer shadow and set all necessary parameters
            OuterShadowEffect Shadow = new OuterShadowEffect();

            Shadow.BlurRadius = 0;
            Shadow.Direction = 50;
            Shadow.Distance = 10;
            Shadow.ColorFormat.Color = Color.LightBlue;

            //shape.TextFrame.TextRange.EffectDag.InnerShadowEffect = Shadow;
            shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow;

            //Save the document
            string result = "SetShadowEffect.pptx";
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