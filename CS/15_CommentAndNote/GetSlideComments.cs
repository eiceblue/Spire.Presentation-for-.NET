using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetSlideComments
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Comments.pptx");

            //Loop through comments
            foreach (ICommentAuthor commentAuthor in presentation.CommentAuthors)
            {
                foreach (Comment comment in commentAuthor.CommentsList)
                {
                    //Get comment information
                    string commentText = comment.Text;
                    string authorName = comment.AuthorName;
                    DateTime time = comment.DateTime;
                    MessageBox.Show("Comment text : "+ comment.Text +"\n"+"Comment author : " + comment.AuthorName + "\n" + "Posted on time : " + comment.DateTime);
                }
            }
        }
    }
}