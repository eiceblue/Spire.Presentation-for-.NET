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
            //Load a PPT document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\FillShapeWithPicture.pptx");

            //Get the first shape and set the style to be Gradient
            IAutoShape shape = ppt.Slides[0].Shapes[0] as IAutoShape;
           
            //Fill the shape with picture
            string picUrl = @"..\..\..\..\..\..\Data\backgroundImg.png";
            shape.Fill.FillType = FillFormatType.Picture;
            shape.Fill.PictureFill.Picture.Url = picUrl;
            shape.Fill.PictureFill.FillType = PictureFillType.Stretch;

            ppt.SaveToFile("FillShapeWithPicture_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("FillShapeWithPicture_result.pptx");
        }
    }
}
