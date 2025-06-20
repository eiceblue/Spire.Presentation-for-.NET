using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddSlideUsingMasterLayout
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

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AppendSlideWithMasterLayout.pptx");

            //Get Master layouts
            ILayout iLayout = presentation.Masters[0].Layouts[0];

            //Append new slide
            presentation.Slides.Append(iLayout);

            //Insert new slide
            presentation.Slides.Insert(1, iLayout);

            //Save to file.
            String result = "output.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
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