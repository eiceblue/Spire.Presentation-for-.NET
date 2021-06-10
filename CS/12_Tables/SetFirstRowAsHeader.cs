using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetFirstRowAsHeader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string loadPath = @"..\..\..\..\..\..\Data\NormalTable.pptx";
            string savePath = "SetFirstRowAsHeader.pptx";
            ITable table = null;

            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(loadPath);

            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = shape as ITable;
                }

            }
            table.FirstRow = true;

            //Save the file
            presentation.SaveToFile(savePath, FileFormat.Pptx2010);
            System.Diagnostics.Process.Start(savePath);
        }
    }
}