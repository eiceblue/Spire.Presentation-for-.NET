using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RotateShape
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document 
            Presentation presentation = new Presentation();

            //append new shape - Triangle
            IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(100, 100, 100, 100));
            //set rotation to 180
            shape.Rotation = 180;

            //set the color and fill style of shape
            shape.Fill.FillType = FillFormatType.Solid;
            shape.Fill.SolidColor.Color = Color.BlueViolet;
            shape.ShapeStyle.LineColor.Color = Color.Black;

            //save the document
            presentation.SaveToFile("RotateShape.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RotateShape.pptx");
        }

        private void lblDescription_Click(object sender, EventArgs e)
        {

        }
    }
}
