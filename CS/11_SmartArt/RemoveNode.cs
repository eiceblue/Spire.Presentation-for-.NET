using Spire.Presentation;
using Spire.Presentation.Diagrams;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveNode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.pptx");

            //get the SmartArt and collect nodes
            ISmartArt sa = presentation.Slides[0].Shapes[0] as ISmartArt;
            ISmartArtNodeCollection nodes = sa.Nodes;

            //remove the node to specific position
            nodes.RemoveNodeByPosition(2);

            presentation.SaveToFile("RemoveNodes.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RemoveNodes.pptx");
        }

        private void lblDescription_Click(object sender, EventArgs e)
        {

        }

        private void pbLogo_Click(object sender, EventArgs e)
        {

        }
    }
}
