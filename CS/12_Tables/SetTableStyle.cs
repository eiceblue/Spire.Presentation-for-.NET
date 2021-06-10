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
            //Creat a ppt document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SetTableStyle.pptx");

            //Get tbe table
            ITable table = null;
            foreach (IShape shape in ppt.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Set the table style from TableStylePreset and apply it to selected table
                    table.StylePreset = TableStylePreset.MediumStyle1Accent2;
                }
            }
            //Save the file
            ppt.SaveToFile("SetTableStyle_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("SetTableStyle_result.pptx");
        }
    }
}
