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
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveNode.pptx");

            //Get the SmartArt and collect nodes
            ISmartArt sa = presentation.Slides[0].Shapes[0] as ISmartArt;
            ISmartArtNodeCollection nodes = sa.Nodes;

            //Remove the node to specific position
            nodes.RemoveNodeByPosition(2);

            presentation.SaveToFile("RemoveNode.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RemoveNode.pptx");
        }
    }
}
