using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace ReplaceTextRetentionStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SomePresentation.pptx");

            presentation.Slides[0].ReplaceFirstText("use", "test", true);

            presentation.Slides[1].ReplaceAllText("Spire", "new spire", true);

            String result = "ReplaceTextRetentionStyle_result.pptx";

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