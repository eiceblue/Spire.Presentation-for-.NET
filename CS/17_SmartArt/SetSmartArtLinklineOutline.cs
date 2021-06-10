using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Diagrams;
using Spire.Presentation.Drawing;

namespace SetSmartArtLinklineOutline
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\CreateSmartArtShape.pptx");
            //Get the specified shape as ISmartArt
            ISmartArt smartArt = ppt.Slides[0].Shapes[0] as ISmartArt;
            int count = smartArt.Nodes.Count;
            ISmartArtNode node;
            //Loop through all smartArts
            for (int i = 0; i < count; i++)
            {
                node = smartArt.Nodes[i];
                //Set the line type
                node.LinkLine.FillType = FillFormatType.Solid;
                //Set the line color
                node.LinkLine.SolidFillColor.Color = Color.Red;
                //Set the line width
                node.LinkLine.Width = 2;
                //Set the line DashStyle
                node.LinkLine.DashStyle = LineDashStyleType.SystemDash;
            }
            //Save the document
            String result = "SetSmartArtLinklineOutline.pptx";
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