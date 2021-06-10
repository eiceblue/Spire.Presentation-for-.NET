using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.Drawing.Printing;

namespace SilentlyPrintPPTByDefaultPrinter
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

            //Print the PowerPoint document to default printer.
            PresentationPrintDocument document = new PresentationPrintDocument(presentation);
            document.PrintController = new StandardPrintController();

            presentation.Print(document);
        }
    }
}