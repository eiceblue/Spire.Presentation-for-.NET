using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetTableStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //creat a ppt document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\test.pptx");
            
            //get tbe table
            ITable table = ppt.Slides[0].Shapes[0] as ITable;

            //set the table style from TableStylePreset and apply it to selected table
            table.StylePreset = TableStylePreset.DarkStyle1Accent6;

            //save the file
            ppt.SaveToFile("tableStyle.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("tableStyle.pptx");
        }
    }
}
