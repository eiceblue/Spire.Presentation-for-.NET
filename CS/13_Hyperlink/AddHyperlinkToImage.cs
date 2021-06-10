using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddHyperlinkToImage
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_5.pptx");

            //Get the first slide.
            ISlide slide = presentation.Slides[0];

            //Add image to slide.
            RectangleF rect = new RectangleF(480, 350, 160, 160);
            IEmbedImage image = slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, @"..\..\..\..\..\..\Data\Logo1.png", rect);

            //Add hyperlink to the image.
            ClickHyperlink hyperlink = new ClickHyperlink("https://www.e-iceblue.com");
            image.Click = hyperlink;

            String result = "Result-AddHyperlinkToImage.pptx";

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