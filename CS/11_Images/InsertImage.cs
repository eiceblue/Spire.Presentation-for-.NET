using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace InsertImage
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\InsertImage.pptx");

            //Insert image to PPT
            string ImageFile2 = @"..\..\..\..\..\..\Data\InsertImage.png";
            RectangleF rect1 = new RectangleF(presentation.SlideSize.Size.Width / 2 - 280, 140, 120, 120);
            IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile2, rect1);
            image.Line.FillType = FillFormatType.None;
    
            //Save the document
            presentation.SaveToFile("InsertImage.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("InsertImage.pptx");
        }
    }
}