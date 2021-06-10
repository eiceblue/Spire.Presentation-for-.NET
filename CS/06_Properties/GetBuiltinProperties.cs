using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;

namespace GetBuiltinProperties
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }
        private void btnRun_Click(object sender, EventArgs e)
        {

            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the PPT document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetProperties.pptx");

            //Get the builtin properties 
            string application = presentation.DocumentProperty.Application;
            string author = presentation.DocumentProperty.Author;
            string company = presentation.DocumentProperty.Company;
            string keywords = presentation.DocumentProperty.Keywords;
            string comments = presentation.DocumentProperty.Comments;
            string category = presentation.DocumentProperty.Category;
            string title = presentation.DocumentProperty.Title;
            string subject = presentation.DocumentProperty.Subject;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();
            content.AppendLine("DocumentProperty.Application: " + application);
            content.AppendLine("DocumentProperty.Author: " + author);
            content.AppendLine("DocumentProperty.Company " + company);
            content.AppendLine("DocumentProperty.Keywords: " + keywords);
            content.AppendLine("DocumentProperty.Comments: " + comments);
            content.AppendLine("DocumentProperty.Category: " + category);
            content.AppendLine("DocumentProperty.Title: " + title);
            content.AppendLine("DocumentProperty.Subject: " + subject);

            //String for .txt file 
            String result = "GetBuiltinProperties_Output.txt";

            //Save them to a txt file
            File.WriteAllText(result, content.ToString());

            //Launching the result file.
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}