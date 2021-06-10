using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ChangeSlidePosition
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

            //Move the first slide to the second slide position
            ISlide slide = presentation.Slides[0];
            slide.SlideNumber = 2;

            //Save the document
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Output.pptx");
        }
    }
}
