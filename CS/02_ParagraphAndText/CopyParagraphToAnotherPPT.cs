using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace CopyParagraphToAnotherPPT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load the source file
            Presentation ppt1 = new Presentation();
            ppt1.LoadFromFile(@"..\..\..\..\..\..\Data\TextTemplate.pptx");

            //Get the text from the first shape on the first slide
            IShape sourceshp = ppt1.Slides[0].Shapes[0];
            string text = ((IAutoShape)sourceshp).TextFrame.Text;

            //Load the target file
            Presentation ppt2 = new Presentation();
            ppt2.LoadFromFile(@"..\..\..\..\..\..\Data\CopyParagraph.pptx");

            //Get the first shape on the first slide from the target file
            IShape destshp = ppt2.Slides[0].Shapes[0];

            //Add the text to the target file
            ((IAutoShape)destshp).TextFrame.Text += "\n\n" + text;

            //Save the document
            string result = "CopyParagraphToAnotherPPT.pptx";
            ppt2.SaveToFile(result, FileFormat.Pptx2013);
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