using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace IsTextBox
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\IsTextboxSample.pptx");

            StringBuilder builder = new StringBuilder();

            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is IAutoShape)
                    {
                        //Judge if the shape is textbox
                        Boolean isTextbox = shape.IsTextBox;
                        builder.AppendLine(isTextbox ? "shape is text box" : "shape is not text box");
                    }
                }
            }

            //Write the content of builder to txt file
            File.WriteAllText("IsTextbox.txt", builder.ToString());
            System.Diagnostics.Process.Start("IsTextbox.txt");
        }
    }
}