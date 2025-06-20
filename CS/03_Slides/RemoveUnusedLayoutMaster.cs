using Spire.Presentation;
using Spire.Presentation.Collections;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveUnusedLayoutMaster
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load document from disk
            Presentation ppt = new Presentation();
            ppt.LoadFromFile("../../../../../../Data/PPTSample_1.pptx");

            //Create an array list
            List<IActiveSlide> list = new List<IActiveSlide>();
            for (int i = 0; i < ppt.Slides.Count; i++)
            {
                //Get the layout used by slide
                IActiveSlide layout = (IActiveSlide)ppt.Slides[i].Layout;
                list.Add(layout);
            }

            //Loop through masters and layouts
            for (int i = 0; i < ppt.Masters.Count; i++)
            {
                IMasterLayouts masterlayouts = ppt.Masters[i].Layouts;
                for (int j = masterlayouts.Count - 1; j >= 0; j--)
                {
                    if (!list.Contains((IActiveSlide)masterlayouts[j]))
                    {
                        //Remove unused layout
                        masterlayouts.RemoveMasterLayout(j);
                    }
                }
            }

            //Save the document
            String outputFile = "RemoveUnusedLayoutMaster_out.pptx";
            ppt.SaveToFile(outputFile, FileFormat.Pptx2013);
            ppt.Dispose();
            System.Diagnostics.Process.Start(outputFile);
        }
    }
}
