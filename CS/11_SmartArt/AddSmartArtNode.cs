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
            //create PPT document
            Presentation presentation = new Presentation();

            //load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.pptx");

            //get the SmartArt
            ISmartArt sa = presentation.Slides[0].Shapes[0] as ISmartArt;

            //add a node
            ISmartArtNode node = sa.Nodes.AddNode();
            //add text and set the text style 
            node.TextFrame.Text = "AddText";
            node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink;


            presentation.SaveToFile("AddNode.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddNode.pptx");
        }
    }
}
