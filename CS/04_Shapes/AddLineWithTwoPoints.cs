using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddLineWithTwoPoints
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

            //Get the first slide
            ISlide slide = ppt.Slides[0];

            //Add line with two points
            IAutoShape line = slide.Shapes.AppendShape(ShapeType.Line, new PointF(50, 50), new PointF(150, 150));
            line.ShapeStyle.LineColor.Color = Color.Red;
            line = slide.Shapes.AppendShape(ShapeType.Line, new PointF(150, 150), new PointF(250, 50));
            line.ShapeStyle.LineColor.Color = Color.Blue;

            //Save the document
            string result = "AddLineWithTwoPoints.pptx";
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