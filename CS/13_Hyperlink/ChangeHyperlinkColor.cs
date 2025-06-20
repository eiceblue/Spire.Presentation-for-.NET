using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace ChangeHyperlinkColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document
            Presentation presentation = new Presentation();

            //Load file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeHyperlinkColor.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Get the theme of the slide
            Theme theme = slide.Theme;

            //Change the color of hyperlink to red
            theme.ColorScheme.HyperlinkColor.Color = Color.Red;
        
            string result = "Result-ChangeHyperlinkColor.pptx";

            //Save to file
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
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