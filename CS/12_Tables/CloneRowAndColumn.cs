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

namespace CloneRowAndColumn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Presentation presentation = new Presentation();
            // Access first slide
            ISlide sld = presentation.Slides[0];

            // Define columns with widths and rows with heights
            double[] widths = { 110, 110, 110 };
            double[] heights = { 50, 30, 30, 30, 30 };

            // Add table shape to slide
            ITable table = presentation.Slides[0].Shapes.AppendTable(presentation.SlideSize.Size.Width / 2 - 275, 90, widths, heights);

            // Add text to the row 1 cell 1
            table[0, 0].TextFrame.Text = "Row 1 Cell 1";

            // Add text to the row 1 cell 2
            table[1, 0].TextFrame.Text = "Row 1 Cell 2";

            // Clone row 1 at end of table
            table.TableRows.Append(table.TableRows[0]);

            // Add text to the row 2 cell 1
            table[0, 1].TextFrame.Text = "Row 2 Cell 1";

            // Add text to the row 2 cell 2
            table[1, 1].TextFrame.Text = "Row 2 Cell 2";

            // Clone row 2 as the 4th row of table
            table.TableRows.Insert(3, table.TableRows[1]);

            //Clone column 1 at end of table
            table.ColumnsList.Add(table.ColumnsList[0]);

            //Clone the 2nd column at 4th column index
            table.ColumnsList.Insert(3, table.ColumnsList[1]);

            string result = "CloneRowAndColumn_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2010);

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