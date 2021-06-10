using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace LinkToLastViewedSlide
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\LastViewedSlide.pptx");
            //Get specified slide
            ISlide slide = ppt.Slides[0];
            //Draw a shape
            IAutoShape autoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 100, 100));
            //Link to last viewed slide show
            autoShape.Click = ClickHyperlink.LastVievedSlide;
            //Save the document
            String result = "GetLastViewedSlide.pptx";
            ppt.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);
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