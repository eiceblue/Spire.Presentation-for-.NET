using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace Alignment
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Alignment.pptx");

            //Get the related shape and set the text alignment
            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[1];
            shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            shape.TextFrame.Paragraphs[1].Alignment = TextAlignmentType.Center;
            shape.TextFrame.Paragraphs[2].Alignment = TextAlignmentType.Right;
            shape.TextFrame.Paragraphs[3].Alignment = TextAlignmentType.Justify;
            shape.TextFrame.Paragraphs[4].Alignment = TextAlignmentType.None;

            //Save the document
            presentation.SaveToFile("alignment_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("alignment_result.pptx");
        }
    }
}