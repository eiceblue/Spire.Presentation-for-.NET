using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace CloneSlideAtTheEnd
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load PPT document from disk
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Append the slide at the end of the document
            presentation.Slides.Append(slide);

            //Save the document
            string result = "ClonePPTAtTheEnd_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}