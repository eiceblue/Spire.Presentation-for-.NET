using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SaveToStream
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PowerPoint file and save it to stream
            Presentation presentation = new Presentation();

            //Set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Append new shape
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, 600, 150));
            shape.Fill.FillType = FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Add text to shape
            shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream.";
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;
            shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 30;

            //Save to Stream
            FileStream to_stream = new FileStream("SaveToStream.pptx", FileMode.Create);
            presentation.SaveToFile(to_stream, FileFormat.Pptx2013);
            to_stream.Close();
            PresentationDocViewer("SaveToStream.pptx");
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