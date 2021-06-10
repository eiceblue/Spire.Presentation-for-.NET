using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Presentation;

namespace OneSlideToSVG
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\OneSlideToSVG.pptx");
          
            //Convert the second slide to SVG
            byte[] svgByte = presentation.Slides[1].SaveToSVG();            
            File.WriteAllBytes("OneSlideToSVG.svg", svgByte);
            //Launch the file
            OutputViewer("OneSlideToSVG.svg");
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }

    }
}