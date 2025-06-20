using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetSlideLayout
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

            //Append a slide and set the layout for slide
            ISlide slide = ppt.Slides.Append(SlideLayoutType.Title);

            //Add content for Title and Text
            IAutoShape shape = slide.Shapes[0] as IAutoShape;
            shape.TextFrame.Text = "Hello Wolrd! ¨C> This is title";

            shape = slide.Shapes[1] as IAutoShape;
            shape.TextFrame.Text = "E-iceblue Support Team -> This is content";

            //Save the document
            string result = "SetSlideLayout.pptx";
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