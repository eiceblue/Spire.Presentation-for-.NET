using Spire.Presentation;
using Spire.Presentation.Charts;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ConvertPdfWithDefaultFont
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load PPT from disk
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertPdfWithDefaultFont.pptx");
            
            //The font is preferred to convert to pdf or pictures, when the font used in the document is not installed in the system
            Presentation.SetDefaultFontName("Arial");
            
            //Save to file
            ppt.SaveToFile("ConvertPdfWithDefaultFont_out.pdf", FileFormat.PDF);

            //Launch and view the resulted file
            PresentationDocViewer("ConvertPdfWithDefaultFont_out.pdf");
        }
        private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
