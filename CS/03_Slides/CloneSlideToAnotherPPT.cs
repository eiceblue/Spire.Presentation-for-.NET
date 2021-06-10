using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CloneSlideToAnotherPPT
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
            Presentation presentation = new Presentation();

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\CloneSlideToAnotherPPT-2.pptx");

            //Load the another document and choose the first slide to be cloned
            Presentation ppt1 = new Presentation();
            ppt1.LoadFromFile(@"..\..\..\..\..\..\Data\CloneSlideToAnotherPPT-1.pptx");
            ISlide slide1 = ppt1.Slides[0];

            //Insert the slide to the specified index in the source presentation
            int index = 1;
            presentation.Slides.Insert(index, slide1); 
 
            //Save the document
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Output.pptx");
        }
    }
}
