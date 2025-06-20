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
using Spire.Presentation.Charts;

namespace HideAudioDuringShow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create Presentation
            Presentation presentation = new Presentation();

            //Load ppt file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\audio.pptx");

            //Get the first slide
            ISlide slide=presentation.Slides[0];

            //Hide Audio during show
            foreach (Shape shape in slide.Shapes)
            {
                if (shape is IAudio)
                {
                    IAudio audio = shape as IAudio;
                    audio.HideAtShowing = true;
                }
            }

            //Save the file
            String result = "HideAudioDuringShow_result.pptx";
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);

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