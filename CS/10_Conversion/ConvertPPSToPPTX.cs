using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ConvertPPSToPPTX
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();
            //Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Conversion.pps");

            //Save the PPS document to PPTX file format
            string result = "ConvertPPSToPPTX.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            //Launch and view the resulted PPTX file
            PresentationDocViewer(result);
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