using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace ExtractComments
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_5.pptx");

            StringBuilder str = new StringBuilder();

            //Get all comments from the first slide.
            Comment[] comments = presentation.Slides[0].Comments;

            //Save the comments in txt file.
            for (int i = 0; i < comments.Length; i++)
            {
                str.Append(comments[i].Text + "\r\n");
            }

            String result = "Result-ExtractComments.txt";

            //Save to file.
            File.WriteAllText(result, str.ToString());

            //Launch the txt file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}