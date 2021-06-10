using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace SplitPPT
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

            for (int i = 0; i < ppt.Slides.Count; i++)
            {
                //Initialize another instance of Presentation, and remove the blank slide
                Presentation newppt = new Presentation();
                newppt.Slides.RemoveAt(0);

                //Append the specified slide from old presentation to the new one
                newppt.Slides.Append(ppt.Slides[i]);

                //Save the document
                string result = string.Format("SplitPPT-{0}.pptx", i);
                newppt.SaveToFile(result, FileFormat.Pptx2010);
                PresentationDocViewer(result);
            }
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