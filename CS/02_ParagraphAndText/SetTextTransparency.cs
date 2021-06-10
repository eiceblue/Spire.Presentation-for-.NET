using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetTextTransparency
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

            //Add a shape
            IAutoShape textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 300, 120));
            textboxShape.ShapeStyle.LineColor.Color = Color.Transparent;
            textboxShape.Fill.FillType = FillFormatType.None;

            //Remove default blank paragraphs
            textboxShape.TextFrame.Paragraphs.Clear();

            //Add three paragraphs, apply color with different alpha values to text
            int alpha = 55;
            for (int i = 0; i < 3; i++)
            {
                textboxShape.TextFrame.Paragraphs.Append(new TextParagraph());
                textboxShape.TextFrame.Paragraphs[i].TextRanges.Append(new TextRange("Text Transparency"));
                textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.FillType = FillFormatType.Solid;
                textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(alpha, Color.Purple);
                alpha += 100;
            }

            //Save the document
            string result = "SetTextTransparency.pptx";
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