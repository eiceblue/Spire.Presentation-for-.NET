using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetOutlineAndEffectForShape
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
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Draw a Rectangle shape
            IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 180, 100, 50));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.SkyBlue;
            //Set outline color
            shape.ShapeStyle.LineColor.Color = Color.Red;
            //Set shadow effect
            PresetShadow shadow = new PresetShadow();
            shadow.ColorFormat.Color = Color.LightSkyBlue;
            shadow.Preset = PresetShadowValue.FrontRightPerspective;
            shadow.Distance = 10.0;
            shadow.Direction = 225.0f;
            shape.EffectDag.PresetShadowEffect = shadow;

            //Draw a Ellipse shape
            shape = slide.Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(400, 150, 100, 100));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.SkyBlue;
            //Set outline color
            shape.ShapeStyle.LineColor.Color = Color.Yellow;
            //Set shadow effect
            GlowEffect glow = new GlowEffect();
            glow.ColorFormat.Color = Color.LightPink;
            glow.Radius = 20.0;
            shape.EffectDag.GlowEffect = glow;

            //Save the document
            string result = "SetOutlineAndEffectForShape.pptx";
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