using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetFormatForLines
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

            //Add a rectangle shape to the slide
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 150, 200, 100));
            //Set the fill color of the rectangle shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.White;
            //Apply some formatting on the line of the rectangle
            shape.Line.Style = TextLineStyle.ThickThin;
            shape.Line.Width = 5;
            shape.Line.DashStyle = LineDashStyleType.Dash;
            //Set the color of the line of the rectangle
            shape.ShapeStyle.LineColor.Color = Color.SkyBlue;

            //Add a ellipse shape to the slide
            shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(400, 150, 200, 100));
            //Set the fill color of the ellipse shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.White;
            //Apply some formatting on the line of the ellipse
            shape.Line.Style = TextLineStyle.ThickBetweenThin;
            shape.Line.Width = 5;
            shape.Line.DashStyle = LineDashStyleType.DashDot;
            //Set the color of the line of the ellipse
            shape.ShapeStyle.LineColor.Color = Color.OrangeRed;

            //Save the document
            string result = "SetFormatForLines.pptx";
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