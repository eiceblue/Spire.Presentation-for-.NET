using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace ApplySlideMaster
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\InputTemplate.pptx");

            //Get the first slide master from the presentation
            IMasterSlide masterSlide = ppt.Masters[0];

            //Customize the background of the slide master
            string backgroundPic = @"..\..\..\..\..\..\Data\bg.png";
            RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
            masterSlide.SlideBackground.Fill.FillType = FillFormatType.Picture;
            IEmbedImage image = masterSlide.Shapes.AppendEmbedImage(ShapeType.Rectangle, backgroundPic, rect);
            masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image as IImageData;

            //Change the color scheme
            masterSlide.Theme.ColorScheme.Accent1.Color = Color.Red;
            masterSlide.Theme.ColorScheme.Accent2.Color = Color.RosyBrown;
            masterSlide.Theme.ColorScheme.Accent3.Color = Color.Ivory;
            masterSlide.Theme.ColorScheme.Accent4.Color = Color.Lavender;
            masterSlide.Theme.ColorScheme.Accent5.Color = Color.Black;

            //Save the document
            string result = "ApplySlideMaster.pptx";
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