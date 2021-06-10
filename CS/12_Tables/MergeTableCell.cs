using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MergeTableCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\MergeTableCell.pptx");

            ITable table = null;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Merge the second row and third row of the first column
                    table.MergeCells(table[0, 1], table[0, 2], false);

                    table.MergeCells(table[3, 4], table[4, 4], true);

                }
            }

			//Save and launch the file
            presentation.SaveToFile("MergeTableCell_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("MergeTableCell_result.pptx");
        }
    }
}
