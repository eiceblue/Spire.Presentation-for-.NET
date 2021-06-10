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

namespace AccessSpecificChildNode
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
            strB.AppendLine("Access SmartArt child node at specific position.");
            strB.AppendLine("Here is the SmartArt child node parameters details:"); 
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ISmartArt)
                {
                    //Get the SmartArt
                    ISmartArt sa = shape as ISmartArt;

                    //Get SmartArt node collection 
                    ISmartArtNodeCollection nodes = sa.Nodes;

                    //Access SmartArt node at index 0
                    ISmartArtNode node = nodes[0];

                    //Access SmartArt child node at index 1
                    ISmartArtNode childNode = node.ChildNodes[1];

                    //Print the SmartArt child node parameters
                    string outString = string.Format("Node text = {0}, Node level = {1}, Node Position = {2}", childNode.TextFrame.Text, childNode.Level, childNode.Position);

                    strB.AppendLine(outString);
                }

            }
            String result = "AccessSpecificChildNode_result.txt";
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