using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveSlide
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\source.pptx");

            //remove the second slide
            presentation.Slides.RemoveAt(1);

            presentation.SaveToFile("RemovedSlide.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RemovedSlide.pptx");
        }
    }
}
