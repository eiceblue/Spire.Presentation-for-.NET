using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.Drawing.Printing;

namespace PrintSpecifiedRangeOfPPTPages
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_6.pptx");

            PresentationPrintDocument document = new PresentationPrintDocument(presentation);

            //Set the document name to display while printing the document. 
            document.DocumentName = "Template_Ppt_6.pptx";

            //Choose to print some pages from the PowerPoint document.
            document.PrinterSettings.PrintRange = PrintRange.SomePages;
            document.PrinterSettings.FromPage = 2;
            document.PrinterSettings.ToPage = 3;

            //Set the number of copies of the document to print.
            document.PrinterSettings.Copies = 2;

            presentation.Print(document);
        }
    }
}