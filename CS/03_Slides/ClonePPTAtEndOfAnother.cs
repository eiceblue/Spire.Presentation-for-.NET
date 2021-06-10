using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ClonePPTAtEndOfAnother
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load source document from disk
            Presentation sourcePPT = new Presentation();
            sourcePPT.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

            //Load destination document from disk
            Presentation destPPT = new Presentation();
            destPPT.LoadFromFile(@"..\..\..\..\..\..\Data\PPTSample_N.pptx");

            //Loop through all slides of source document
            foreach (ISlide slide in sourcePPT.Slides)
            {
                //Append the slide at the end of destination document
                destPPT.Slides.Append(slide);
            }

            //Save the document
            string result = "ClonePPTAtEndOfAnother_result.pptx";
            destPPT.SaveToFile(result, FileFormat.Pptx2013);

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