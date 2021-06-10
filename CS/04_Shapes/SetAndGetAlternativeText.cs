using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetAndGetAlternativeText
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ShapeTemplate.pptx");

            //Get the first slide
            ISlide slide = ppt.Slides[0];

            //Set the alternative text (title and description)
            slide.Shapes[0].AlternativeTitle = "Rectangle";
            slide.Shapes[0].AlternativeText = "This is a Rectangle";

            //Get the alternative text (title and description)
            string alternativeText = null;
            string title = slide.Shapes[0].AlternativeTitle;
            alternativeText += "Title: " + title + "\r\n";
            string description = slide.Shapes[0].AlternativeText;
            alternativeText += "Description: " + description;

            //Save the document
            string result = "SetAlternativeText.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);

            //Save the alternative text to Text file
            result = "GetAlternativeText.txt";
            File.WriteAllText(result, alternativeText);
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