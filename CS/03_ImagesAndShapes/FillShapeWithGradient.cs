using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FillShapeWithGradient
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation
            Presentation ppt = new Presentation();
            //Add a rectangle to the slide
            IAutoShape GradientShape = (IAutoShape)ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(200, 100, 287, 100));
            
            //Set the Fill Type of the Shape and color to Gradient
            GradientShape.Fill.FillType = FillFormatType.Gradient;
            GradientShape.Fill.Gradient.GradientStops.Append(0, Color.Purple);
            GradientShape.Fill.Gradient.GradientStops.Append(1, Color.Red);
            
            ppt.SaveToFile("CreateGradientShape.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("CreateGradientShape.pptx");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
