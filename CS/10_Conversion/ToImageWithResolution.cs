using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToImageWithResolution
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

            // Render the first slide to an image with custom DPI
            Image img = ppt.SaveToImage(0,100,100);

            //Save image to file
            string result = "ToImageWithResolution.png";
            img.Save(result, System.Drawing.Imaging.ImageFormat.Png);

            //Launch the PPT file
            System.Diagnostics.Process.Start(result);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}