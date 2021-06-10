using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;

namespace ChangeNodeText
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

            //Load the PPT
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddSmartArtNode.pptx");

            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ISmartArt)
                {
                    //Get the SmartArt and collect nodes
                    ISmartArt smartArt = shape as ISmartArt;
                    //Obtain the reference of a node by using its Index  
                    // select second root node
                    ISmartArtNode node = smartArt.Nodes[1]; 
                    // Set the text of the TextFrame 
                    node.TextFrame.Text = "Second root node";
                }
            }
            String result = "ChangeNodeText_result.pptx";
            //Save the file
            presentation.SaveToFile(result, FileFormat.Pptx2010);

            Viewer(result);
        }

        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}