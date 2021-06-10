using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace AddImageWatermark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_1.pptx");

            //Get the image you want to add as image watermark.
            IImageData image = presentation.Images.Append(Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png"));

            //Set the properties of SlideBackground, and then fill the image as watermark.
            presentation.Slides[0].SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom;
            presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Picture;
            presentation.Slides[0].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch;
            presentation.Slides[0].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image;

            String result = "Result-AddImageWatermark.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}