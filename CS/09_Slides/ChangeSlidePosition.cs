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
            //create PPT document
            Presentation presentation = new Presentation();

            //load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ReorderSlidePosition.pptx");

            //move the first slide to the second slide position
            ISlide slide = presentation.Slides[0];
            slide.SlideNumber = 2;

            //save the document
            presentation.SaveToFile("ChangedPosition.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("ChangedPosition.pptx");
        }
    }
}
