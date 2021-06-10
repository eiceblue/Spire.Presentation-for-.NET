using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetTextMargins
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

            //Append a new shape
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, 450, 150));

            //Set margins for text inside shapes
            shape.Fill.FillType = FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.LightBlue;
            shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify;
            shape.TextFrame.Text = "Using Spire.Presentation, developers will find an easy and effective method to create, read, write, modify, convert and print PowerPoint files on any .Net platform. It's worthwhile for you to try this amazing product.";
            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Rounded MT Bold");
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;

            //Set the margins for the text frame
            shape.TextFrame.MarginTop = 10;
            shape.TextFrame.MarginBottom = 35;
            shape.TextFrame.MarginLeft = 15;
            shape.TextFrame.MarginRight = 30;

            //Save the document
            string result = "SetTextMargins.pptx";
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