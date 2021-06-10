using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace AddSection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\BlankSample.pptx");

            //Get the second slide
            ISlide slide = ppt.Slides[1];

            //Append section with section name at the end
            ppt.SectionList.Append("E-iceblue01");
            //Add section with slide
            ppt.SectionList.Add("section1", slide);

            String result = "AddSection.pptx";
            ppt.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);
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