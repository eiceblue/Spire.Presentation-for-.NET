using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Diagrams;
using System.IO;

namespace ExtractTextFromSmartArt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractTextFromSmartArt.pptx");

            //Traverse through all the slides of the PPT file and find the SmartArt shapes.
            StringBuilder st = new StringBuilder();
           st.AppendLine("Below is extracted text from SmartArt:");
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                for (int j = 0; j < presentation.Slides[i].Shapes.Count; j++)
                {
                    if (presentation.Slides[i].Shapes[j] is ISmartArt)
                    {
                        ISmartArt smartArt = presentation.Slides[i].Shapes[j] as ISmartArt;

                        //Extract text from SmartArt and append to the StringBuilder object.
                        for (int k = 0; k < smartArt.Nodes.Count; k++)
                        {
                            st.AppendLine(smartArt.Nodes[k].TextFrame.Text);
                        }
                    }
                }
            }

            String result = "Result-ExtractTextFromSmartArt.txt";

            //Save to file.
            File.WriteAllText(result, st.ToString());

            //Launch the file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}