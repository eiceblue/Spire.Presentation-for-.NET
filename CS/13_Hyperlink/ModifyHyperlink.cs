
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ModifyHyperlink
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

            //Find the hyperlinks you want to edit.
            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];

            //Edit the link text and the target URL.
            shape.TextFrame.TextRange.ClickAction.Address = "http://www.e-iceblue.com";
            shape.TextFrame.TextRange.Text = "E-iceblue";

            String result = "Result-ModifyHyperlink.pptx";

            //Save to file.
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