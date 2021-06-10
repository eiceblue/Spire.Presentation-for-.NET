using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExtractText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractText.pptx");
            
            StringBuilder sb = new StringBuilder();
            //Foreach the slide and extract text
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is IAutoShape)
                    {
                        foreach (TextParagraph tp in (shape as IAutoShape).TextFrame.Paragraphs)
                        {
                            sb.Append(tp.Text + Environment.NewLine);
                        }
                    }

                }
            }
            File.WriteAllText("ExtractText_result.txt", sb.ToString());
            System.Diagnostics.Process.Start("ExtractText_result.txt");
        }
    }
}
