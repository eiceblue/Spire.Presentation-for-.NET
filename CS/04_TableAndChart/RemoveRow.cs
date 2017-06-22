using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveRow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\table.pptx");

            //get the table in PPT document
            ITable table = null;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //remove the second column
                    table.ColumnsList.RemoveAt(1, false);

                    //remove the second row
                    table.TableRows.RemoveAt(1, false);
                }
            }
            //save the document
            presentation.SaveToFile("RemoveRow.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RemoveRow.pptx");
        }

        private void lblDescription_Click(object sender, EventArgs e)
        {

        }

        private void pbLogo_Click(object sender, EventArgs e)
        {

        }
    }
}
