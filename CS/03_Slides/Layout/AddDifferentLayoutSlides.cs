using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddDifferentLayoutSlides
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

            //Remove the default slide
            presentation.Slides.RemoveAt(0);

            //Loop through slide layouts
            foreach (SlideLayoutType type in Enum.GetValues(typeof(SlideLayoutType)))
            {
                //Append slide by specifing slide layout
                presentation.Slides.Append(type);
            }

            //Save the document
            string result = "AddDifferentLayoutSlides_result.pptx";
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