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

namespace AccessChildNodes
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SmartArt.pptx");

            StringBuilder strB = new StringBuilder();
            strB.AppendLine("Access SmartArt child nodes.");
            strB.AppendLine("Here is the SmartArt child node parameters details:"); 
            string outString = "";
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ISmartArt)
                {
                    //Get the SmartArt and collect nodes
                    ISmartArt sa = shape as ISmartArt;
                    ISmartArtNodeCollection nodes = sa.Nodes;
                 
                    int position = 0;
                    //Access the parent node at position 0
                    ISmartArtNode node = nodes[position];
                    ISmartArtNode childnode;
                    //Traverse through all child nodes inside SmartArt
                    for (int i = 0; i < node.ChildNodes.Count; i++)
                    {
                        //Access SmartArt child node at index i
                        childnode = node.ChildNodes[i];
                        //Print the SmartArt child node parameters                       
                        outString = string.Format("Node text = {0}, Node level = {1}, Node Position = {2}", childnode.TextFrame.Text, childnode.Level, childnode.Position);
                        strB.AppendLine(outString);
                    }
 
                }

            }
            String result = "AccessChildNode_result.txt";
            //Save the file
            File.WriteAllText(result, strB.ToString());

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