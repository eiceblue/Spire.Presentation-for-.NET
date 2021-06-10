using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace MergeSelectedSlides
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

            //Remove the first slide
            ppt.Slides.RemoveAt(0);

            //Load two PPT files
            Presentation ppt1 = new Presentation(@"..\..\..\..\..\..\Data\InputTemplate.pptx", FileFormat.Pptx2013);
            Presentation ppt2 = new Presentation(@"..\..\..\..\..\..\Data\TextTemplate.pptx", FileFormat.Pptx2013);
            
            //Append all slides in ppt1 to ppt
            for (int i = 0; i < ppt1.Slides.Count; i++)
            {
                ppt.Slides.Append(ppt1.Slides[i]);
            }

            //Append the second slide in ppt2 to ppt
            ppt.Slides.Append(ppt2.Slides[1]);

            //Save the document
            string result = "MergeSelectedSlides.pptx";
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