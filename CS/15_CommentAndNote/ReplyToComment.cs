using Spire.Presentation;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace ReplyToComment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create ppt file
            Presentation ppt = new Presentation();        

            //Create Comment author
            ICommentAuthor author = ppt.CommentAuthors.AddAuthor("E-iceblue", "comment");

            //Add comment
            ppt.Slides[0].AddComment(author, "Add comment", new System.Drawing.Point(18, 25), DateTime.Now);
            Comment comment = ppt.Slides[0].Comments[0];

            //Add reply to Comment
            if (!comment.IsReply)
            {
                comment.Reply(author, "Add Reply1", DateTime.Now);
                comment.Reply(author, "Add Reply2", DateTime.Now);
            }

            //delete first reply
            ppt.Slides[0].DeleteComment(author, "Add Reply1");

            //Save the result ppt file
            ppt.SaveToFile(@"AddReplyToComment.pptx", FileFormat.Pptx2013);
            System.Diagnostics.Process.Start("AddReplyToComment.pptx");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}