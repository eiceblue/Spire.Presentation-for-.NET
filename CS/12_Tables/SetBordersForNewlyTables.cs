using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetBordersForNewlyTables
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

            //Set the table width and height for each table cell.
            double[] tableWidth = new double[] { 100, 100, 100, 100, 100 };
            double[] tableHeight = new double[] { 20, 20 };

            //Traverse all the border type of the table.
            foreach (TableBorderType item in Enum.GetValues(typeof(TableBorderType)))
            {
              //Add a table to the presentation slide with the setting width and height
                ITable itable = presentation.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight);

                //Add some text to the table cell.
                itable.TableRows[0][0].TextFrame.Text = "Row";
                itable.TableRows[1][0].TextFrame.Text = "Column";

                //Set the border type, border width and the border color for the table.
                itable.SetTableBorder(item, 1.5, Color.Red);
            }

            String result = "Result-SetBordersForNewlyTables.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}