using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveRowsAndColumns
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveRowsAndColumns.pptx");

            //Get the table in PPT document
            ITable table = null;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Remove the second column
                    table.ColumnsList.RemoveAt(1, false);

                    //Remove the second row
                    table.TableRows.RemoveAt(1, false);
                }
            }
            //Save and launch the document
            presentation.SaveToFile("RemoveRowsAndColumns_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RemoveRowsAndColumns_result.pptx");
        }
    }
}
