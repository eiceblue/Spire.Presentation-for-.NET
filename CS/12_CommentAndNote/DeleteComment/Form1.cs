using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DeleteComment
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\DeleteComment.pptx");

            //Replace the text in the comment
            presentation.Slides[0].Comments[1].Text = "Replace comment";

            //Delete the third comment
            presentation.Slides[0].DeleteComment(presentation.Slides[0].Comments[2]);

            //Save the document
            presentation.SaveToFile("DeleteComment.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("DeleteComment.pptx");
        }
    }
}
