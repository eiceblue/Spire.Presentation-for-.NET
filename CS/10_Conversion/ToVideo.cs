using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToVideo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation ppt = new Presentation();
            //Load the PPT document from disk.
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ToVideo.pptx");

            // Download ffmpeg before converting PPT to video
            // Windows: Download ffmpeg-x.x.x-essentials_build.zip file and extract it.
            // CentOS: Install via command line. Note that a lower ffmpeg version may cause some slide transition effects to fail.
            SaveToVideoOption saveOption = new SaveToVideoOption(@"..\..\..\..\..\..\Data\ffmpeg-8.1.1-essentials_build\ffmpeg-8.1.1-essentials_build\bin");
            // Set frame rate
            saveOption.Fps = 10;
            // Set duration for each slide
            saveOption.DurationForEachSlide = 5;
            ppt.SaveToVideoOption = saveOption;

            String result = "toVideo.mp4";
            ppt.SaveToFile(result, FileFormat.MP4);
            ppt.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}