using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToHTML
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

            //Save the document to HTML format
            string result = "ToHTML.html";
            ppt.SaveToFile(result, FileFormat.Html);
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