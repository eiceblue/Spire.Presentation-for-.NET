using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Diagrams;
using Spire.Presentation.Drawing;

namespace SetSmartArtNodeOutline
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
            //Set ISmartArt form special shape
            ISmartArt smartArt = ppt.Slides[0].Shapes[0] as ISmartArt;
            int count = smartArt.Nodes.Count;
            ISmartArtNode node;
            //Loop through all nodes
            for (int i = 0; i < count; i++)
            {
                node = smartArt.Nodes[i];
                //Set the fill format type
                node.Line.FillType = FillFormatType.Solid;
                //Set the line style
                node.Line.Style = TextLineStyle.ThinThin;
                //Set the line color
                node.Line.SolidFillColor.Color = Color.Red;
                //Set the line width
                node.Line.Width = 2;
            }
            //Save the document
            String result = @"SetSmartArtNodeOutline.pptx";
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