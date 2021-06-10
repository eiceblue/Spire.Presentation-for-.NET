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

namespace AccessSmartArt
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
            strB.AppendLine("Access SmartArt nodes.");
            strB.AppendLine("Here is the SmartArt node parameters details:"); 
            string outString="";
            ISmartArtNode node;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ISmartArt)
                {
                    //Get the SmartArt
                    ISmartArt sa = shape as ISmartArt;
                    
                    ISmartArtNodeCollection nodes = sa.Nodes;
               
                    //Traverse through all nodes inside SmartArt
                    for (int i = 0; i < nodes.Count; i++)
                    {
                        //Access SmartArt node at index i
                        node = nodes[i];
                        //Print the SmartArt node parameters
                        outString = string.Format("Node text = {0}, Node level = {1}, Node Position = {2}", node.TextFrame.Text, node.Level, node.Position);
                        strB.AppendLine(outString);
                    }
                }

            }
            String result = "AccessSmartArt_result.txt";
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