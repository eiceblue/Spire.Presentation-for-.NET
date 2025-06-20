using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;
using Spire.Presentation.Charts;

namespace ConvertPPTToOFD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create Presentation
            Presentation presentation = new Presentation();

            //Load ppt file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\CopyParagraph.pptx");

            //Save the PPT document to OFD format
            String result = "ConvertPPTToOFD_result.ofd";
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.OFD);

            //Launching the result file.
            Viewer(result);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}