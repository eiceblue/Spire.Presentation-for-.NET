using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetImageTransparency
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

            //Create an Image from the specified file
            string imagePath = @"..\..\..\..\..\..\Data\Logo.png";
            Image image = Image.FromFile(imagePath);
            float width = image.Width;
            float height = image.Height;
            RectangleF rect1 = new RectangleF(200, 100, width, height);
            //Add a shape
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect1);
            shape.Line.FillType = FillFormatType.None;
            //Fill shape with image
            shape.Fill.FillType = FillFormatType.Picture;
            shape.Fill.PictureFill.Picture.Url = imagePath;
            shape.Fill.PictureFill.FillType = PictureFillType.Stretch;
            //Set transparency on image
            shape.Fill.PictureFill.Picture.Transparency = 50;

            //Save the document
            string result = "SetImageTransparency.pptx";
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