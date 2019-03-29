using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddComment
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddComment.pptx");

            //Comment author
            ICommentAuthor author = presentation.CommentAuthors.AddAuthor("E-iceblue", "comment:");

            //Add comment
            presentation.Slides[0].AddComment(author, "Add comment", new System.Drawing.PointF(18, 25), DateTime.Now);

            //Save the document
            presentation.SaveToFile("AddComment.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddComment.pptx");
        }
    }
}
