using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace CreateSlideMasterAndApply
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

            ppt.SlideSize.Type = SlideSizeType.Screen16x9;

            //Add slides
            for (int i = 0; i < 4; i++)
            {
                ppt.Slides.Append();
            }

            //Get the first default slide master
            IMasterSlide first_master = ppt.Masters[0];

            //Append another slide master
            ppt.Masters.AppendSlide(first_master);
            IMasterSlide second_master = ppt.Masters[1];

            //Set different background image for the two slide masters
            string pic1 = @"..\..\..\..\..\..\Data\bg.png";
            string pic2 = @"..\..\..\..\..\..\Data\Setbackground.png";
            //The first slide master
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            first_master.SlideBackground.Fill.FillType = FillFormatType.Picture;
            IEmbedImage image1 = first_master.Shapes.AppendEmbedImage(ShapeType.Rectangle, pic1, rect);
            first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image1 as IImageData;
            //The second slide master
            second_master.SlideBackground.Fill.FillType = FillFormatType.Picture;
            IEmbedImage image2 = second_master.Shapes.AppendEmbedImage(ShapeType.Rectangle, pic2, rect);
            second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image2 as IImageData;

            //Apply the first master with layout to the first slide
            ppt.Slides[0].Layout = first_master.Layouts[1];

            //Apply the second master with layout to other slides
            for (int i = 1; i < ppt.Slides.Count; i++)
            {
                ppt.Slides[i].Layout = second_master.Layouts[8];
            }

            //Save the document
            string result = "CreateSlideMasterAndApply.pptx";
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