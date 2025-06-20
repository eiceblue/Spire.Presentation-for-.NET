using Spire.Presentation;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace GetSlideText
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
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\GetSlideText.pptx");

            //Foreach the slide and get text
            foreach (ISlide slide in ppt.Slides)
            {
                ArrayList arrayList = slide.GetAllTextFrame();
                foreach (String text in arrayList)
                {
                    MessageBox.Show(text);
                }
            }
        }
    }
}
