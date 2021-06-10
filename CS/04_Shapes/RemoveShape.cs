using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace RemoveShape
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

            //Load doucment from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\FindShapeByAltText.pptx");

            //Loop through slides
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                ISlide slide = presentation.Slides[i];
                //Loop through shapes
                for (int j = 0; j < slide.Shapes.Count; j++)
                {
                    IShape shape = slide.Shapes[j];
                    //Find the shapes whose alternative text contain "Shape"
                    if(shape.AlternativeText.Contains("Shape"))
                    {
                        slide.Shapes.Remove(shape);
                        j--;
                    }
                }
            }

            //Save the document
            string result = "RemoveShape_result.pptx";
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