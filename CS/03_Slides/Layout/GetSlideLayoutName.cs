using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace GetSlideLayoutName
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
			Presentation presentation=new Presentation();

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\GetSlideLayoutName.pptx");

            StringBuilder builder = new StringBuilder();

            //Loop through the slides of PPT document
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                //Get the name of slide layout
                string name = presentation.Slides[i].Layout.Name;
                builder.AppendLine(String.Format("The name of slide {0} layout is: {1}", i,name));
            }

            //Save the document
            string result="GetSlideLayoutName_out.txt";
            File.WriteAllText(result, builder.ToString());

            //Launch the Pdf file
            DocumentViewer(result);
        }
        private void DocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }

        }
    }
}
