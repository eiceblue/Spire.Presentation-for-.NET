using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetImageFrameFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Load an image
            string imageFile = @"..\..\..\..\..\..\Data\iceblueLogo.png";
            Image image = Image.FromFile(imageFile);

            //Add the image in document
            IImageData imageData = presentation.Images.Append(image);
            RectangleF rect = new RectangleF(100,100,imageData.Width/2,imageData.Height/2);
            IEmbedImage pptImage = presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, imageData, rect);

            //Set the formatting of the image frame
            pptImage.Line.FillFormat.FillType = FillFormatType.Solid;
            pptImage.Line.FillFormat.SolidFillColor.Color = Color.LightBlue;
            pptImage.Line.Width = 5;
            pptImage.Rotation = -45;

            //Save the document
            string result = "SetImageFrameFormat_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}