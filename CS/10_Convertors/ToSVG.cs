using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using System.IO;

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

            Queue<byte[]> svgBytes=presentation.SaveToSVG();
            for (int i = 0; i < svgBytes.Count; i++)
            {
                FileStream fs = new FileStream(String.Format("{0}.svg", i), FileMode.Create);
                byte[] bt = svgBytes.Dequeue();
                fs.Write(bt, 0, bt.Length);
            }
            
        }

    }
}