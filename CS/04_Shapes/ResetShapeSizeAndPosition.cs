using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ResetShapeSizeAndPosition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();
            //Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ShapeTemplate.pptx");

            //Define the original slide size
            float currentHeight = ppt.SlideSize.Size.Height;
            float currentWidth = ppt.SlideSize.Size.Width;

            //Change the slide size as A3
            ppt.SlideSize.Type = SlideSizeType.A3;

            //Define the new slide size
            float newHeight = ppt.SlideSize.Size.Height;
            float newWidth = ppt.SlideSize.Size.Width;

            //Define the ratio from the old and new slide size
            float ratioHeight = newHeight / currentHeight;
            float ratioWidth = newWidth / currentWidth;

            //Reset the size and position of the shape on the slide
            foreach (ISlide slide in ppt.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    shape.Height = shape.Height * ratioHeight;
                    shape.Width = shape.Width * ratioWidth;

                    shape.Left = shape.Left * ratioHeight;
                    shape.Top = shape.Top * ratioWidth;
                }
            }

            //Save the document
            string result = "ResetShapeSizeAndPosition.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
		}
	
		private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}