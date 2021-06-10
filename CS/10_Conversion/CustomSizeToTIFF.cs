using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;

namespace CustomSizeToTIFF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the original PPT document from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Indent.pptx");

            //Get the first slide
            ISlide slide= presentation.Slides[0];

            //Create a new PPT document
            Presentation newPresentation = new Presentation();

            //Remove the default slide 
            newPresentation.Slides.RemoveAt(0);

            //Define a new size
            SizeF size = new SizeF(200F, 200F);

            //Set PPT slide size
            newPresentation.SlideSize.Size = size;

            //Insert the slide of original PPT
            newPresentation.Slides.Insert(0, slide);
    
            //String for output file 
            String result = "Output1.tiff";

            //Save the second slide to PDF
            newPresentation.SaveToFile(result, Spire.Presentation.FileFormat.Tiff);

        
            //Launching the result file.
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}