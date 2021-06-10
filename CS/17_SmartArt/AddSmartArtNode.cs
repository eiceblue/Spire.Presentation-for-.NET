using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Diagrams;

namespace AddSmartArtNode
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
            Presentation presentation = new Presentation();

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddSmartArtNode.pptx");

            //Get the SmartArt
            ISmartArt sa = presentation.Slides[0].Shapes[0] as ISmartArt;

            //Add a node
            ISmartArtNode node = sa.Nodes.AddNode();
            //Add text and set the text style 
            node.TextFrame.Text = "AddText";
            node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink;

            presentation.SaveToFile("AddSmartArtNode.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddSmartArtNode.pptx");
        }
    }
}
