using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SaveChartAsImage
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

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SaveChartAsImage.pptx");

            //Save chart as image in .png format
            Image image = presentation.Slides[0].Shapes.SaveAsImage(0);
            image.Save("Chart_result.png", System.Drawing.Imaging.ImageFormat.Png);
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            System.IO.Stream stream = ppt.Slides[0].SaveAsImage();
            byte[] buff = new byte[stream.Length];
            stream.Read(buff, 0, buff.Length);
            FileStream fs = new FileStream("Chart_result.png", FileMode.Create);
            fs.Write(buff);
            fs.Close();
			*/
            
            System.Diagnostics.Process.Start("Chart_result.png");
        }
    }
}
