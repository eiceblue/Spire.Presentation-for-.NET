using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace AddShapes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Set background Image
            string ImageFile = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
            presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

            //Append new shape - Triangle and set style
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(115, 130, 100, 100));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.LightGreen;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Append new shape - Ellipse
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(290, 130, 150, 100));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.LightSkyBlue;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Append new shape - Heart
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Heart, new RectangleF(470, 130, 130, 100));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.Red;
            shape.ShapeStyle.LineColor.Color = Color.LightGray;


            //Append new shape - FivePointedStar
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.FivePointedStar, new RectangleF(90, 270, 150, 150));
            shape.Fill.FillType = FillFormatType.Gradient;
            shape.Fill.SolidColor.Color = Color.Black;
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Append new shape - Rectangle
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(320, 290, 100, 120));
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.Pink;
            shape.ShapeStyle.LineColor.Color = Color.LightGray;

            //Append new shape - BentUpArrow
            shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.BentUpArrow, new RectangleF(470, 300, 150, 100));

            //Set the color of shape
            shape.Fill.FillType = FillFormatType.Gradient;
            shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.Olive);
            shape.Fill.Gradient.GradientStops.Append(0, KnownColors.PowderBlue);
            shape.ShapeStyle.LineColor.Color = Color.White;

            //Save the document
            presentation.SaveToFile("AddShapes_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddShapes_result.pptx");
        }
    }
}