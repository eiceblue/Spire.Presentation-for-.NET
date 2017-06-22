using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;

namespace Spire.Presentation.Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\edit.pptx");

            //edit the first shape
            IAutoShape shape = (IAutoShape) presentation.Slides[0].Shapes[0];
            TextParagraph para = new TextParagraph();
            para.Text = "Edit Sample";
            para.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
            para.TextRanges[0].FontHeight = 24;
            shape.TextFrame.Paragraphs.Append(para);

            //save the document
            presentation.SaveToFile("edited.pptx", FileFormat.Pptx2007);
            System.Diagnostics.Process.Start("edited.pptx");
        }
    }
}