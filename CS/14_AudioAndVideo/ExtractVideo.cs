using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;

namespace ExtractVideo
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
            Presentation presentation = new Presentation();

            //Load the PPT document from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\video.pptx");

            //Define a variable 
            int i = 0;

            //String for output file 
            String result = string.Format(@"Video{0}.avi", i);

            //Traverse all the slides of PPT file
            foreach (ISlide slide in presentation.Slides)
            {
                //Traverse all the shapes of slides
                foreach (IShape shape in slide.Shapes)
                {
                    //If shape is IVideo
                    if (shape is IVideo)
                    {
                        //Save the video
                        (shape as IVideo).EmbeddedVideoData.SaveToFile(result);
                        i++;
                    }
                }
            }
            //Launching the result file.
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}