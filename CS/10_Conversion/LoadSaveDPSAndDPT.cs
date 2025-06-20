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

namespace LoadSaveDPSAndDPT
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

            //Load .dps or .dpt file
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\sample.dps",Spire.Presentation.FileFormat.Dps);
             //presentation.LoadFromFile(@"..\..\..\..\..\..\Data\sample.dpt",Spire.Presentation.FileFormat.Dpt);

            //Save the .dps or .dpt file
            String result = "LoadSaveDPSAndDPT_result.dps";
            presentation.SaveToFile(result, Spire.Presentation.FileFormat.Dps);
            //presentation.SaveToFile("LoadSaveDPSAndDPT_result.dpt", Spire.Presentation.FileFormat.Dpt);
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