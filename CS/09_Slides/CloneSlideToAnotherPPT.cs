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
            //create PPT document and load PPT file from disk
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\source.pptx");

            //Load the another document and choose the first slide to be cloned.
            Presentation ppt1 = new Presentation();
            ppt1.LoadFromFile(@"..\..\..\..\..\..\Data\Presentation1.pptx");
            ISlide slide1 = ppt1.Slides[0];

            //Insert the slide to the specified index in the source presentation
            int index = 1;
            presentation.Slides.Insert(index, slide1); 
 
            //save the document
            presentation.SaveToFile("ClonedSlide.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("ClonedSlide.pptx");
        }
    }
}
