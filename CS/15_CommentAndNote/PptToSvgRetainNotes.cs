using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace PptToSvgRetainNotes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_5.pptx");

            //Retain the notes while converting PowerPoint file to svg file.
            presentation.IsNoteRetained = true;

            //Convert presentation slides to svg file.
            Queue<byte[]> bytes = presentation.SaveToSVG();

            int length = bytes.Count;
            for (int i = 0; i < length; i++)
            {
                String result = string.Format(@"output_{0}.svg", i);
                FileStream filestream = new FileStream(result, FileMode.Create);
                byte[] outputBytes = bytes.Dequeue();
                filestream.Write(outputBytes, 0, outputBytes.Length);

                //Launch the PowerPoint file.
                PptDocumentViewer(result);
            }

            presentation.Dispose();
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}