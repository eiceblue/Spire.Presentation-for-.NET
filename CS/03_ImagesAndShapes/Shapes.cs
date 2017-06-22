using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;

namespace Spire.Presentation.Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //append new shape - Triangle
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(50, 100, 100, 100));
            
            //set the color and fill style of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.LightGreen;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //append new shape - Ellipse
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(270, 100, 150, 100));
            shape.ShapeStyle.LineColor.Color = Color.White;

            //append new shape - FivePointedStar
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.FivePointedStar, new RectangleF(50, 270, 150, 150));
            
            //set the color of shape
            shape.Fill.FillType = FillFormatType.Gradient;
            shape.Fill.SolidColor.Color = Color.Black;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //append new shape - Rectangle
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(300, 300, 100, 120));
            
            //set the color of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.Tomato;
            shape.ShapeStyle.LineColor.Color = Color.Tomato;

            //append new shape - BentUpArrow
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.BentUpArrow, new RectangleF(500, 300, 150, 100));

            //set the color of shape
            shape.Fill.FillType = FillFormatType.Gradient;
            shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.Olive);
            shape.Fill.Gradient.GradientStops.Append(0, KnownColors.PowderBlue);
            shape.ShapeStyle.LineColor.Color = Color.White;

            //save the document
            presentation.SaveToFile("shape.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("shape.pptx");
        }
    }
}