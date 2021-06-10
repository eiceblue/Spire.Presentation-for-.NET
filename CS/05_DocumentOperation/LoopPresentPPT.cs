using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace LoopPresentPPT
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

            //Set the Boolean value of ShowLoop as true
            ppt.ShowLoop = true;

            //Set the PowerPoint document to show animation and narration
            ppt.ShowAnimation = true;
            ppt.ShowNarration = true;
            //Use slide transition timings to advance slide
            ppt.UseTimings = true;

            //Save the document
            string result = "LoopPresentPPT.pptx";
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