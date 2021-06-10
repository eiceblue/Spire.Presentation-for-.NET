using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace UngroupShapes
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\GroupShapes.pptx");
            //Get the GroupShape
            GroupShape groupShape = ppt.Slides[0].Shapes[0] as GroupShape;
            //Ungroup the shapes
            ppt.Slides[0].Ungroup(groupShape);
            //Save the document
            String result = "UngroupShapes.pptx";
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