using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation.Collections;
using Spire.Presentation;

namespace Indent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Indent.pptx");

            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];
            ParagraphCollection paras = shape.TextFrame.Paragraphs;

            //Set the paragraph style for first paragraph
            paras[0].Indent = 20;
            paras[0].LeftMargin = 10;
            paras[0].SpaceAfter = 10;
           
            //Set the paragraph style of the third paragraph 
            paras[2].Indent = -100;
            paras[2].LeftMargin = 40;
            paras[2].SpaceBefore = 0;
            paras[2].SpaceAfter = 0;
          
            //Save the document
            presentation.SaveToFile("Indent_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("Indent_result.pptx");
        }
    }
}