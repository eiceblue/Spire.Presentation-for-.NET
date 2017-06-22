using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CreateSmartArtShape
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create a document and add a smartart
            Presentation pres = new Presentation();
            Spire.Presentation.Diagrams.ISmartArt sa = pres.Slides[0].Shapes.AppendSmartArt(20, 40, 300, 300, Spire.Presentation.Diagrams.SmartArtLayoutType.Gear);

            //set type and color of smartart
            sa.Style = Spire.Presentation.Diagrams.SmartArtStyleType.SubtleEffect;
            sa.ColorStyle = Spire.Presentation.Diagrams.SmartArtColorType.GradientLoopAccent3;

            //remove all shapes
            foreach (object a in sa.Nodes)
                sa.Nodes.RemoveNode(0);

            //add two custom shapes with text
            Spire.Presentation.Diagrams.ISmartArtNode node = sa.Nodes.AddNode();
            sa.Nodes[0].TextFrame.Text = "aa";
            node = sa.Nodes.AddNode();
            node.TextFrame.Text = "bb";
            node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Black;

            //save and launch the file
            pres.SaveToFile("SmartArt.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("SmartArt.pptx");
        }
    }
}
