using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetGlobalCustomFonts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Set global custom fonts 
            Presentation.SetCustomFontsDirctory(@"..\..\..\..\..\..\Data\fonts");

            //Create a PPT document
            Presentation ppt = new Presentation();

            //Load PPT file 
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\toPDF.pptx");

            //Save the PPT to PDF file format
            String result = "output.pdf";
            ppt.SaveToFile(result, FileFormat.PDF);

            System.Diagnostics.Process.Start(result);
        }
    }
}