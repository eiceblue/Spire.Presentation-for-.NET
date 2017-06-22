using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FillShapeWithPicture
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
            Presentation ppt = new Presentation();

            //add a rectangle to the slide
            IAutoShape shape = (IAutoShape)ppt.Slides[0].Shapes.AppendShape(ShapeType.DoubleWave, new RectangleF(100, 100, 400, 200));
           
            //fill the shape with picture
            string picUrl = @"..\..\..\..\..\..\Data\bg.png";
            shape.Fill.FillType = FillFormatType.Picture;
            shape.Fill.PictureFill.Picture.Url = picUrl;
            shape.Fill.PictureFill.FillType = PictureFillType.Stretch;
            shape.ShapeStyle.LineColor.Color = Color.Transparent;

            ppt.SaveToFile("FillShapeWithPicture.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("FillShapeWithPicture.pptx");
        }
    }
}
