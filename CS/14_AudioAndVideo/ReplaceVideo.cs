using Spire.Presentation;
using Spire.Presentation.Collections;
using System;
using System.IO;
using System.Windows.Forms;


namespace ReplaceVideo
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\video.pptx");

            VideoCollection videos = ppt.Videos;

            //Traverse all the slides of PPT file
            foreach (ISlide sld in ppt.Slides)
            {
                //Traverse all the shapes of slides
                foreach (Shape sp in sld.Shapes)
                {
                    //If shape is IVideo
                    if (sp is IVideo)
                    {
                        //Replace the video
                        IVideo video = sp as IVideo;
                        //Load the video document from disk.
                        byte[] bts = File.ReadAllBytes(@"..\..\..\..\..\..\Data\repleaceVido.mp4");
                        VideoData videoData = videos.Append(bts);
                        video.EmbeddedVideoData = videoData;
                    }
                }
            }

            //Save the document
            string outputFile = "replaceVideo.pptx";
            ppt.SaveToFile(outputFile, FileFormat.Pptx2013);

            //Launch the PPT file
            FileViewer(outputFile);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
