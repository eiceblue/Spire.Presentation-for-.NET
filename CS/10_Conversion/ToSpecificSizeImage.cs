using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToSpecificSizeImage
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

            //Save the first slide to Image and set the image size to 600*400
            Image img = ppt.Slides[0].SaveAsImage(600, 400);           
            
            //Save image to file
            string result = "ToSpecificSizeImage.png";
            img.Save(result, System.Drawing.Imaging.ImageFormat.Png);
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            using (var images = ppt.Slides[0].SaveAsImage(600, 400))
            {
                FileStream fileStream = new FileStream("ToSpecificSizeImage.png", FileMode.Create, FileAccess.Write);
                images.CopyTo(fileStream);
                fileStream.Flush();
                images.Dispose();
            }
            */
            
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