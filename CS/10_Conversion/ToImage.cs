using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create PPT document
            Presentation presentation = new Presentation();

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToImage.pptx");

            //Save PPT document to images
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                String fileName = String.Format("ToImage-img-{0}.png", i);
                Image image = presentation.Slides[i].SaveAsImage();
                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                System.Diagnostics.Process.Start(fileName);
            }

        }
    }
}