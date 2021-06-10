using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ExtractImageFromSpecificSlide
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Images.pptx");

            //Get the pictures on the second slide and save them to image file
            int i = 0;
            //Traverse all shapes in the second slide
            foreach (IShape s in ppt.Slides[1].Shapes)
            {
                //It is the SlidePicture object
                if (s is SlidePicture)
                {
                    //Save to image
                    SlidePicture ps = s as SlidePicture;
                    ps.PictureFill.Picture.EmbedImage.Image.Save(string.Format("{0}.png", i));
                    i++;
                }
                //It is the PictureShape object
                if (s is PictureShape)
                {
                    //Save to image
                    PictureShape ps = s as PictureShape;
                    ps.EmbedImage.Image.Save(string.Format("{0}.png", i));
                    i++;
                }
            }
		}
    }
}