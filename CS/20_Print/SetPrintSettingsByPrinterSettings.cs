using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.Drawing.Printing;

namespace SetPrintSettingsByPrinterSettings
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

            //Use PrinterSettings object to print presentation slides.
            PrinterSettings ps = new PrinterSettings();
            ps.PrintRange = PrintRange.AllPages;
            ps.PrintToFile = true;
            String result = "Result-SetPrintSettingsByPrinterSettingsObject.xps";
            ps.PrintFileName = (result);

            //Print the slide with frame.
            presentation.SlideFrameForPrint = true;

            //Print the slide with Grayscale.
            presentation.GrayLevelForPrint = true;

            //Print 4 slides horizontal.
            presentation.SlideCountPerPageForPrint = PageSlideCount.Four;
            presentation.OrderForPrint = Order.Horizontal;

            //Only select some slides to print.
            //presentation.SelectSlidesForPrint("1", "3");

            //Print the document.
            presentation.Print(ps);

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