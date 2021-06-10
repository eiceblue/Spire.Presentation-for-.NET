using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetPrintSettingsByPrintDocument
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

            //Use PrintDocument object to print presentation slides.
            PresentationPrintDocument document = new PresentationPrintDocument(presentation);

            //Print document to virtual printer.
            document.PrinterSettings.PrinterName = "Microsoft XPS Document Writer";

            //Print the slide with frame.
            presentation.SlideFrameForPrint = true;

            //Print 4 slides horizontal.
            presentation.SlideCountPerPageForPrint = PageSlideCount.Four;
            presentation.OrderForPrint = Order.Horizontal;

            //Print the slide with Grayscale.
            presentation.GrayLevelForPrint = true;

            //Set the print document name.          
            document.DocumentName = "Template_Ppt_6.pptx";

            document.PrinterSettings.PrintToFile = true;
            String result = "Result-SetPrintSettingsByPrintDocumentObject.xps";
            document.PrinterSettings.PrintFileName = (result);

            presentation.Print(document);

            //Launch the file.
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