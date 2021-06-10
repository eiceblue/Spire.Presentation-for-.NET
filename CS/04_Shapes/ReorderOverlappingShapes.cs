using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ReorderOverlappingShapes
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\OverlappingShapes.pptx");

            //Get the first shape of the first slide
            IShape shape = ppt.Slides[0].Shapes[0];
            //Change the shape's zorder
            ppt.Slides[0].Shapes.ZOrder(1, shape);

            //Save the document
            string result = "ReorderOverlappingShapes.pptx";
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