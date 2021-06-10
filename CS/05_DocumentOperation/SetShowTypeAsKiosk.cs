using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetShowTypeAsKiosk
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\InputTemplate.pptx");

            //Specify the presentation show type as kiosk
            ppt.ShowType = SlideShowType.Kiosk;

            //Save the document
            string result = "SetShowTypeAsKiosk.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
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