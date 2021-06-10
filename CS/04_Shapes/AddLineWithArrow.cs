using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddLineWithArrow
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

            //Add a line to the slides and set its color to red
            IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, new RectangleF(150, 100, 100, 100));
            shape.ShapeStyle.LineColor.Color = Color.Red;
            //Set the line end type as StealthArrow
            shape.Line.LineEndType = LineEndType.StealthArrow;

            //Add a line to the slides and use default color
            shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, new RectangleF(300, 150, 100, 100));
            shape.Rotation = -45;
            //Set the line end type as TriangleArrowHead
            shape.Line.LineEndType = LineEndType.TriangleArrowHead;

            //Add a line to the slides and set its color to Green
            shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, new RectangleF(450, 100, 100, 100));
            shape.ShapeStyle.LineColor.Color = Color.Green;
            shape.Rotation = 90;
            //Set the line begin type as TriangleArrowHead
            shape.Line.LineBeginType = LineEndType.StealthArrow;

            //Save the document
            string result = "AddLineWithArrow.pptx";
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