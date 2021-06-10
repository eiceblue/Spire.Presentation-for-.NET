using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace CopyShapesBetweenSlides
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load the sample document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\CopyShapesBetweenSlides.pptx");

            //Define the source slide and target slide
            ISlide sourceSlide = ppt.Slides[0];
            ISlide targetSlide = ppt.Slides[1];

            //Copy the first shape from the source slide to the target slide
            targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[0]);

            string result = "CopyShapesBetweenSlides-result.pptx";
            //Save the document to file 
            ppt.SaveToFile(result, FileFormat.Pptx2013);
          
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}