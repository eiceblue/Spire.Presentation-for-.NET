using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace HideShape
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

            //Load document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\FindShapeByAltText.pptx");

            //Loop through slides
            foreach (ISlide slide in presentation.Slides)
            {
                //Loop through shapes in the slide
                foreach (IShape shape in slide.Shapes)
                {
                    //Find the shape whose alternative text is Shape1
                    if (shape.AlternativeText.CompareTo("Shape1") == 0)
                    {
                        //Hide the shape
                        shape.IsHidden = true;
                    }
                }
            }

            //Save the document
            string result = "HideShape_result.pptx";
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