using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace LoadFromStream
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

            //Load PowerPoint file from stream
            FileStream from_stream = File.OpenRead(@"..\..\..\..\..\..\Data\InputTemplate.pptx");
            ppt.LoadFromStream(from_stream, FileFormat.Pptx2013);        

            //Save the document
            string result = "LoadFromStream.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            from_stream.Dispose();
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