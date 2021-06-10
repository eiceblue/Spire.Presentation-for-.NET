using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace HideSlide
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load PPT file from disk
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\HideSlide.pptx");

            //Hide the second slide
            ppt.Slides[1].Hidden = true;

            //Save the document
            ppt.SaveToFile("HideSlide.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("HideSlide.pptx");
        }
    }
}
