using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace Print
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Print.pptx");

            //Print
            PrinterSettings printerSettings = new PrinterSettings();
            printerSettings.FromPage = 0;
            printerSettings.ToPage = presentation.Slides.Count-1;
            presentation.Print(printerSettings);
        }
    }
}