using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace BordersAndShading
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
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\BordersAndShading.pptx");

            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];

            //Set line color and width of the border
            shape.Line.FillType = FillFormatType.Solid;
            shape.Line.Width = 3;
            shape.Line.SolidFillColor.Color = Color.LightYellow;

            //Set the gradient fill color of shape
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Gradient;
            shape.Fill.Gradient.GradientShape = Spire.Presentation.Drawing.GradientShapeType.Linear;
            shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.LightBlue);
            shape.Fill.Gradient.GradientStops.Append(0, KnownColors.LightSkyBlue);

            //Set the shadow for the shape
            Spire.Presentation.Drawing.OuterShadowEffect shadow = new Spire.Presentation.Drawing.OuterShadowEffect();
            shadow.BlurRadius = 20;
            shadow.Direction = 30;
            shadow.Distance = 8;
            shadow.ColorFormat.Color = Color.LightSeaGreen;
            shape.EffectDag.OuterShadowEffect = shadow;
           
            //Save the document
            presentation.SaveToFile("BordersAndShading_result.pptx", FileFormat.Pptx2007);
            System.Diagnostics.Process.Start("BordersAndShading_result.pptx");
        }
    }
}