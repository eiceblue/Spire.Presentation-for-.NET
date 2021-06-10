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
            //Load a PPT document
            Presentation ppt = new Presentation();

            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\FillShapeWithGradient.pptx");

            //Get the first shape and set the style to be Gradient
            IAutoShape GradientShape = ppt.Slides[0].Shapes[0] as IAutoShape;
            GradientShape.Fill.FillType = FillFormatType.Gradient;
            GradientShape.Fill.Gradient.GradientStops.Append(0, Color.LightSkyBlue);
            GradientShape.Fill.Gradient.GradientStops.Append(1, Color.LightGray);

            ppt.SaveToFile("FillShapeWithGradient_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("FillShapeWithGradient_result.pptx");
        }

    }
}
