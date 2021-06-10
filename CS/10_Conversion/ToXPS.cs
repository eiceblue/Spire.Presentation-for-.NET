using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToXPS
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Conversion.pptx");

            //Save the the XPS file
            string result = "ToXPS.xps";
            ppt.SaveToFile(result, FileFormat.XPS);
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