using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ClonePPTMasterToAnother
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load PPT1 from disk
            Presentation presentation1 = new Presentation();
            presentation1.LoadFromFile(@"..\..\..\..\..\..\Data\CloneMaster1.pptx");

            //Load PPT2 from disk
            Presentation presentation2 = new Presentation();
            presentation2.LoadFromFile(@"..\..\..\..\..\..\Data\CloneMaster2.pptx");

            //Add masters from PPT1 to PPT2
            foreach (IMasterSlide masterSlide in presentation1.Masters)
            {
                presentation2.Masters.AppendSlide(masterSlide);
            }
            
            //Save the document
            string result = "ClonePPTMasterToAnother.pptx";
            presentation2.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
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