using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace GroupShapes
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
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Create two shapes in the slide
            IShape rectangle = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 180, 200, 40));
            rectangle.Fill.FillType = FillFormatType.Solid;
            rectangle.Fill.SolidColor.KnownColor = KnownColors.SkyBlue;
            rectangle.Line.Width = 0.1f;
            IShape ribbon = slide.Shapes.AppendShape(ShapeType.Ribbon2, new RectangleF(290, 155, 120, 80));
            ribbon.Fill.FillType = FillFormatType.Solid;
            ribbon.Fill.SolidColor.KnownColor = KnownColors.LightPink;
            ribbon.Line.Width = 0.1f;

            //Add the two shape objects to an array list
            ArrayList list = new ArrayList();
            list.Add(rectangle);
            list.Add(ribbon);

            //Group the shapes in the list
            ppt.Slides[0].GroupShapes(list);

            //Save the document
            string result = "GroupShapes.pptx";
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