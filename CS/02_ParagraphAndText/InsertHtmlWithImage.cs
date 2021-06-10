using System;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Collections;

namespace InsertHtmlWithImage
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
            ShapeList shapes = ppt.Slides[0].Shapes;

            shapes.AddFromHtml("<html><div><p>First paragraph</p><p><img src='..\\..\\..\\..\\..\\..\\Data\\Logo.png'/></p><p>Second paragraph </p></html>");

            //Save the document
            string result = "InsertHtmlWithImage.pptx";
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