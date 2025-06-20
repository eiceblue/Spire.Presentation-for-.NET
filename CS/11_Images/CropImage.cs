using Spire.Presentation;
using System;
using System.Windows.Forms;

namespace CropImage
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\CropImage.pptx");

            //Get the first shape in first slide
            IShape shape = ppt.Slides[0].Shapes[0];

            //If the shape is SlidePicture
            if (shape is SlidePicture)
            {
                SlidePicture slidePicture = (SlidePicture)shape;
                //Crop image
                slidePicture.Crop(slidePicture.Left + 50f, slidePicture.Top + 50f, 100f, 200f);
            }

            //Save the document
            String result = "CropImage_out.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
        }

        private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
