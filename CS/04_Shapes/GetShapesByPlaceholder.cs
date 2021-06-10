using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace GetShapesByPlaceholder
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
            Presentation ppt = new Presentation();
            //Load the document from disk
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\GetShapesByPlaceholder.pptx");
            //Get Placeholder
            Placeholder placeholder = ppt.Slides[1].Shapes[0].Placeholder;
            //Get Shapes by Placeholder
            IShape[] shapes = ppt.Slides[1].GetPlaceholderShapes(placeholder);
            string text = "";
            //Iterate over all the shapes
            for (int i = 0; i < shapes.Length; i++)
            {
                //If shape is IAutoShape
                if (shapes[i] is IAutoShape)
                {
                    IAutoShape autoShape = shapes[i] as IAutoShape;
                    if (autoShape.TextFrame != null)
                    {
                        text += autoShape.TextFrame.Text + "\r\n";
                        
                    }
                }
            }
            String result = "GetShapesByPlaceholder_output.txt";
            File.WriteAllText(result, text);

            //Launch the PowerPoint file
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