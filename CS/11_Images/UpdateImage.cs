using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace UpdateImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();
            //Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\UpdateImage.pptx");

            //Get the first slide
            ISlide slide = ppt.Slides[0];

            //Append a new image to replace an existing image
            IImageData image = ppt.Images.Append(Image.FromFile(@"..\..\..\..\..\..\Data\PresentationIcon.png"));

			//////////////////Use the following code for netstandard dlls/////////////////////////
			/*
			FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\PresentationIcon.png", FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = new byte[fileStream.Length];
            fileStream.Read(bytes, 0, bytes.Length);
            fileStream.Close();
            Stream stream = new MemoryStream(bytes);
            IImageData image = ppt.Images.Append(stream);
            stream.Close();
            fileStream.Close();
			*/
            
            //Replace the image which title is "image1" with the new image
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is SlidePicture)
                {
                    if (shape.AlternativeTitle == "image1")
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image;
                    }
                }
            }

            //Save the document
            string result = "UpdateImage.pptx";
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