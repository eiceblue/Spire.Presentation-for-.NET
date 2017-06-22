using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Spire.Presentation.Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create PPT document
            Presentation presentation = new Presentation();

            //load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\source.pptx");

            //save PPT document to images
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                String fileName = String.Format("result-img-{0}.png", i);
                Image image = presentation.Slides[i].SaveAsImage();
                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                System.Diagnostics.Process.Start(fileName);
            }

        }
    }
}