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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveSlide.pptx");

            //Remove slide by index
            presentation.Slides.RemoveAt(0);

            //Remove slide by its reference
            ISlide slide = presentation.Slides[1];
            presentation.Slides.Remove(slide);

            //Save the document
            string result = "RemoveSlide_result.pptx";
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
