using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetMasterBackground
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Set the slide background of master
            presentation.Masters[0].SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom;
            presentation.Masters[0].SlideBackground.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            presentation.Masters[0].SlideBackground.Fill.SolidColor.Color = Color.LightSalmon;

            //Save the document
            string result = "SetSlideMasterBackground_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

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