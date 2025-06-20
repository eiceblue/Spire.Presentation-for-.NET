using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetAdvanceAfterTime
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
            Presentation ppt = new Presentation();

            //Load the document from disk
            ppt.LoadFromFile(@"..\..\..\..\..\..\..\Data\SetTransitions.pptx");
   
            //Traverse all slides
            for (int i = 0; i < ppt.Slides.Count; i++)
            {
                ppt.Slides[i].SlideShowTransition.AdvanceOnClick = true;

                //Set the time
                ppt.Slides[i].SlideShowTransition.AdvanceAfterTime = 5000;
            }

            string result = "Result.pptx";
            //Save the document
            ppt.SaveToFile(result, FileFormat.Pptx2010);
            
            PresentationDocViewer(result);
        }

        private static void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}