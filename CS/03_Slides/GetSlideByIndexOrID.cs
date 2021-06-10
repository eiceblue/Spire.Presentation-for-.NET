using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetSlideByIndexOrID
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\BlankSample_N.pptx");

            //Get slide by index 0
            ISlide slide1 = presentation.Slides[0];
            //Append a shape in the slide
            IAutoShape shape1=slide1.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 200, 100));
            //Add text in the shape
            shape1.TextFrame.Text = "Get slide by index";

            //Get slide by slide ID
            ISlide slide2 = presentation.FindSlide((int)presentation.Slides[1].SlideID);
            //Append a shape in the slide
            IAutoShape shape2 = slide2.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 200, 100));
            //Add text in the shape
            shape2.TextFrame.Text = "Get slide by slide id";

            //Save the document
            string result = "GetSlideByIndexOrID_result.pptx";
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