using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace CloneSlideWithinAPPT
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

            //Get a list of slides and choose the first slide to be cloned
            ISlide slide = ppt.Slides[0];

            //Insert the desired slide to the specified index in the same presentation
            int index = 1;
            ppt.Slides.Insert(index, slide);

            //Save the document
            string result = "CloneSlideWithinAPPT.pptx";
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