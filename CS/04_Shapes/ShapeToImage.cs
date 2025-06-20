﻿using Spire.Presentation;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace ShapeToImage
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ShapeToImage.pptx");

            for (int i = 0; i < presentation.Slides[0].Shapes.Count; i++)
            {
                string fileName = String.Format("Picture-{0}.png", i);
                //Save shapes as images
                Image image = presentation.Slides[0].Shapes[i].SaveAsImage();

                //The following method also can save shape as image
                //Image image = presentation.Slides[0].Shapes.SaveAsImage(i);

                //Write image to Png
                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                System.Diagnostics.Process.Start(fileName);
            }
        }
    }
}
