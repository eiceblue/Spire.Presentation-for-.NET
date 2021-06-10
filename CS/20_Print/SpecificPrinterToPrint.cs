using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;
using System.Drawing.Printing;

namespace SpecificPrinterToPrint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load the PPT document from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

            //New PrintSeetings
            PrinterSettings printerSettings = new PrinterSettings();

            //Set landscape for page
            printerSettings.DefaultPageSettings.Landscape = true;

            //Specific the printer
            printerSettings.PrinterName = "Microsoft XPS Document Writer";

            //Print 
            presentation.Print(printerSettings);
        }
    }
}