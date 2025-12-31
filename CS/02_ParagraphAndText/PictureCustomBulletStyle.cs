using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace PictureCustomBulletStyle
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Bullets.pptx");

            //Get the second shape on the first slide
            IAutoShape shape = ppt.Slides[0].Shapes[1] as IAutoShape;

            //Traverse through the paragraphs in the shape
            foreach (TextParagraph paragraph in shape.TextFrame.Paragraphs)
            {
                //Set the bullet style of paragraph as picture
                paragraph.BulletType = TextBulletType.Picture;
                
                //////////////////Use the following code for netstandard dlls/////////////////////////
                /*
                FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\icon.png", FileMode.Open, FileAccess.Read, FileShare.Read);
                byte[] bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, bytes.Length);
                fileStream.Close();
                Stream stream = new MemoryStream(bytes);
                paragraph.BulletPicture.EmbedImage = ppt.Images.Append(stream);
                stream.Close();
                */
                
                //Load a picture
                Image bulletPicture = Image.FromFile(@"..\..\..\..\..\..\Data\icon.png");
                //Add the picture as the bullet style of paragraph
                paragraph.BulletPicture.EmbedImage = ppt.Images.Append(bulletPicture);
            }

            //Save the document
            string result = "PictureCustomBulletStyle.pptx";
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