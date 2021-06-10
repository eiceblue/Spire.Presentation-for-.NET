using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.TimeLine;

namespace ObtainSoundEffect
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Animation.pptx");

            //Get the first slide
            ISlide slide = ppt.Slides[0];

            //Get the audio in a time node
            TimeNodeAudio audio = slide.Timeline.MainSequence[0].TimeNodeAudios[0];

            //Get the properties of the audio, such as sound name, volume or detect if it's mute
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("SoundName: " + audio.SoundName);
            sb.AppendLine("Volume: " + audio.Volume);
            sb.AppendLine("IsMute: " + audio.IsMute);

            //Save the properties of the audio to Text file
            string result = "ObtainSoundEffect.txt";
            File.WriteAllText(result, sb.ToString());
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