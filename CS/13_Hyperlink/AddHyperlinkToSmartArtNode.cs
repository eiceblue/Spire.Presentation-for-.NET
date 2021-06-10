using System;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Diagrams;

namespace AddHyperlinkToSmartArtNode
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArtNode.pptx");

            //Get the smartArt shape
            ISmartArt sr = ppt.Slides[0].Shapes[0] as ISmartArt;
            //Add hylerlinks to the nodes
            ISmartArtNode node = sr.Nodes[0];
            node.Click = new ClickHyperlink(ppt.Slides[1]);
            node = sr.Nodes[1];
            node.Click = new ClickHyperlink(ppt.Slides[2]);
            node = sr.Nodes[2];
            node.Click = new ClickHyperlink(ppt.Slides[3]);
            //Save the document
            string result = "AddHyperlinkToSmartArtNode.pptx";
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