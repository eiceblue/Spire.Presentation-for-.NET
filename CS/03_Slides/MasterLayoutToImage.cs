using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace MasterLayoutToImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load the file 
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\CloneMaster2.pptx");

            // Iterate the masters
            for (int i = 0; i < ppt.Masters[0].Layouts.Count; i++)
            {
                // Save layouts as images
                Image image = ppt.Masters[0].Layouts[i].SaveAsImage();
                String fileName = String.Format("{0}.png", i);
                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
            }

            ppt.Dispose();            
        }
    }
}