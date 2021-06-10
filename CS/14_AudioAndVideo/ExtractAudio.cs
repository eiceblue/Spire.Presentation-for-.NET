using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ExtractAudio
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string loadPath = @"..\..\..\..\..\..\Data\audio.pptx";
            string outPath = @"extrctAudio.wav";
            byte[] AudioData = null;

            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(loadPath);

            foreach (Shape shape in presentation.Slides[0].Shapes)
            {
                if (shape is IAudio)
                {
                    IAudio audio = shape as IAudio;
                    AudioData = audio.Data.Data;
                }
            }

            using (FileStream fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(AudioData, 0, AudioData.Length);

            }
            System.Diagnostics.Process.Start(outPath);
        }
    }
}