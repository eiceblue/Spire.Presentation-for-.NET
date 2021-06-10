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

namespace PrintMultipleSlidesIntoOnePage
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
            Presentation ppt = new Presentation();
            //Load the document from disk
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\PrintMultipleSlidesIntoOnePage.pptx");
            PresentationPrintDocument document = new PresentationPrintDocument(ppt);
            
            //Set print task name
            document.DocumentName = "print task 1";
            document.PrintOrder = Order.Horizontal;
            document.SlideFrameForPrint = true;

            //Set Gray level when printing
            document.GrayLevelForPrint = true;
            //Set four slides on one page
            document.SlideCountPerPageForPrint = PageSlideCount.Four;
         
            //Set continuous print area
            document.PrinterSettings.PrintRange = PrintRange.SomePages;
            document.PrinterSettings.FromPage = 1;
            document.PrinterSettings.ToPage = ppt.Slides.Count - 1;

            //Set discontinuous print area
            //document.SelectSldiesForPrint("1", "2-4");

            ppt.Print(document);
            ppt.Dispose();
        }
    }
}