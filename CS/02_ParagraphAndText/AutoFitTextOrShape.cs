using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AutoFitTextOrShape
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

            //Set the AutofitType property to Shape
            IAutoShape textShape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 100, 150, 80));
            textShape2.TextFrame.Text = "Resize shape to fit text.";
            textShape2.TextFrame.AutofitType = TextAutofitType.Shape;

            //Set the AutofitType property to Normal
            IAutoShape textShape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(400, 100, 150, 80));
            textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape.";
            textShape1.TextFrame.AutofitType = TextAutofitType.Normal;

            //Save the document
            string result = "AutoFitTextOrShape.pptx";
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