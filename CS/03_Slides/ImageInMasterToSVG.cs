using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ImageInMasterToSVG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\Data\ImageInMasterToSVG.pptx");

            //Get the master collection
            IMasterSlide masterSlide = presentation.Masters[0];

            int num = 1;
            for (int i = 0; i < masterSlide.Shapes.Count; i++)
            {
                IShape shape = masterSlide.Shapes[i];
                if (shape is SlidePicture)
                {
                    SlidePicture ps = shape as SlidePicture;
                    byte[] svgByte = ps.SaveAsSvgInSlide();
                    FileStream fs = new FileStream( num + ".svg", FileMode.Create);
                    fs.Write(svgByte, 0, svgByte.Length);
                    fs.Close();
                    num++;
                }
            }
        }
    }
}
