using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Spire.Presentation;

namespace CompressImages
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractImage.pptx");

            foreach (ISlide slide in presentation.Slides)
            {
                foreach (Spire.Presentation.IShape shape in slide.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        SlidePicture ps = shape as SlidePicture;
                        // Compress image: Remove cropped areas (true) and compress image resolution to 150 DPI 
                        ps.PictureFill.CompressImage(true, 50f);
                    }
                }
            }
            //Save to file
            String result = "CompressImages_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PPT file
            System.Diagnostics.Process.Start(result);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}