using System;
using System.IO;
using System.Windows.Forms;
using Spire.Presentation;

namespace SaveShapeAsSvgInSlide
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_7.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            // Iterate the shapes in the slide
            for (int i=0;i<slide.Shapes.Count;i++)
            {
                // Save the shapes
                byte[] svgByte = slide.Shapes[i].SaveAsSvgInSlide();
                FileStream fs = new FileStream("shapePath_" + i + ".svg", FileMode.Create);

                // Close the stream
                fs.Write(svgByte, 0, svgByte.Length);
                fs.Close();
            }

            // Dispose the Presentation object
            presentation.Dispose();
        }
    }
}