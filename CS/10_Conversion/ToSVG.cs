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

namespace ToSVG
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToSVG.pptx");

            //Retain note when converting a PPT document to SVG files
            presentation.IsNoteRetained = true;

            Queue<byte[]> svgBytes = presentation.SaveToSVG();
            int count = svgBytes.Count;
            for (int i = 0; i < count; i++)
            {
                byte[] bt = svgBytes.Dequeue();
                String fileName = String.Format("ToSVG-{0}.svg", i);
                FileStream fs = new FileStream(fileName, FileMode.Create);
                fs.Write(bt, 0, bt.Length);
                System.Diagnostics.Process.Start(fileName);
            }
        }
    }
}