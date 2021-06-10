using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetAnimationForAnimateText
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\..\Data\Animation.pptx");

            //Set the AnimateType as Letter
            ppt.Slides[0].Timeline.MainSequence[0].IterateType = Spire.Presentation.Drawing.TimeLine.AnimateType.Letter;

            //Set the IterateTimeValue for the animate text
            ppt.Slides[0].Timeline.MainSequence[0].IterateTimeValue = 10;

            //Save the document
            string result = "SetAnimationForAnimateText.pptx";
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