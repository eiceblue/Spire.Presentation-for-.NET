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

namespace AssistantNode
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
            ISmartArtNode node;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ISmartArt)
                {
                    //Get the SmartArt and collect nodes
                    ISmartArt smartArt = shape as ISmartArt;
                   
                    ISmartArtNodeCollection nodes = smartArt.Nodes;
               
                    //Traverse through all nodes inside SmartArt
                    for (int i = 0; i < nodes.Count; i++)
                    {
                        //Access SmartArt node at index i
                        node = nodes[i];
                        // Check if node is assitant node
                        if (!node.IsAssistant)
                        {
                            //Set node as assitant node
                            node.IsAssistant = true;
                        }
                    }
                }
            }
            String result = "AssistantNode_result.pptx";
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