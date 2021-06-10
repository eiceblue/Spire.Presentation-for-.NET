using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace GetShapeGroupAltText
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

            //Load document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetShapeGroupAltText.pptx");

            StringBuilder builder=new StringBuilder();

            //Loop through slides and shapes
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is GroupShape)
                    {
                        //Find the shape group
                        GroupShape groupShape = shape as GroupShape;
                        foreach (IShape gShape in groupShape.Shapes)
                        {
                            //Append the alternative text in builder
                            builder.AppendLine(gShape.AlternativeText);
                        }
                    }
                }
            }

            //Write the content in txt file
            string output="GetShapeAltText_result.txt";
            File.WriteAllText(output, builder.ToString());

            //Launch the txt file
            OutputViewer(output);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}