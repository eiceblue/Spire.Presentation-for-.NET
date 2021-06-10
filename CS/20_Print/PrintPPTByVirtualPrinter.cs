using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace PrintPPTByVirtualPrinter
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

            //Print PowerPoint document to virtual printer (Microsoft XPS Document Writer).
            PresentationPrintDocument document = new PresentationPrintDocument(presentation);
            document.PrinterSettings.PrinterName = "Microsoft XPS Document Writer";

            presentation.Print(document);
        }
    }
}