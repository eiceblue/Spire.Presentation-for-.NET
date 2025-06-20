using System;
using System.IO;
using System.Windows.Forms;
using Spire.Presentation;

namespace LoadEncryptedStream
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a Presentation instance
            Presentation ppt = new Presentation();

            //Load PowerPoint file from stream
            FileStream from_stream = File.OpenRead(@"..\..\..\..\..\..\Data\\OpenEncryptedPPT.pptx");

            // The password
            String password = "123456";

            // Load the encrypted stream with the provided password
            ppt.LoadFromStream(from_stream, FileFormat.Auto, password);

            // Save the decrypted document to disk
            ppt.SaveToFile("output/result.pptx", FileFormat.Pptx2013);

            // Dispose the Presentation object
            ppt.Dispose();
        }
    }
}