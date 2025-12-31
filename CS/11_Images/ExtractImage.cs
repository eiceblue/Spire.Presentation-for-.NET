using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExtractImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractImage.pptx");

            for (int i = 0; i < ppt.Images.Count; i++)
            {
                string ImageName = string.Format("Images-{0}.png", i);
                //Extract image
                Image image = ppt.Images[i].Image;
                image.Save(ImageName);
              
                //////////////////Use the following code for netstandard dlls/////////////////////////
                /*
                SkiaSharp.SKImage image = SkiaSharp.SKImage.FromBitmap(ppt.Images[i].Image);
                FileStream fileStream = new FileStream(ImageName, FileMode.Create, FileAccess.Write);
                image.Encode().SaveTo(fileStream);
                fileStream.Flush();
                image.Dispose();
                */          
            }
            ppt.Dispose();
        }
    }
}
