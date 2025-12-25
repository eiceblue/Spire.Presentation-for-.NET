using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace ToSvgUsingGraphicTag
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractImage.pptx");
            //When saving a PPT document to SVG, save the graphics in the PPT document as image tags
            presentation.SaveToSvgOption.ConvertPictureUsingGraphicTag = true;
            for (int i = 0; i <presentation.Slides.Count; i++)
            {       
                String fileName = String.Format("ToSVG-{0}.svg", i);
                FileStream fs = new FileStream(fileName, FileMode.Create);
                //Convert the  slide to SVG
                byte[] silde = presentation.Slides[i].SaveToSVG();
                fs.Write(silde, 0, silde.Length);
                System.Diagnostics.Process.Start(fileName);
            }
          
        }
    }
}
