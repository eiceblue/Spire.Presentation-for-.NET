using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;

namespace LockAspectRatio
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

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Table.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            StringBuilder str = new StringBuilder();
            foreach (IShape shape in slide.Shapes)
            {
                //Verify if it is table
                if (shape is ITable)
                {
                    ITable table = (ITable)shape;
                    //Lock aspect ratio
                    table.ShapeLocking.AspectRatioProtection = true;
                }
            }

            string result = "LockAspectRatio_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);
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