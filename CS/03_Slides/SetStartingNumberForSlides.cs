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

namespace SetStartingNumberForSlides
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

            //Load the PPT document from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

            //Set 5 as the starting number
            presentation.FirstSlideNumber = 5;

            //String for output file 
            String result = "Output.pptx";

            //Save file
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013);

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