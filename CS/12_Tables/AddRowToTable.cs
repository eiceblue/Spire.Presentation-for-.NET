using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddRowToTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_1.pptx");

            //Get the table within the PowerPoint document.
            ITable table = presentation.Slides[0].Shapes[0] as ITable;

            //Get the second row.
            TableRow row = table.TableRows[1];

            //Clone the row and add it to the end of table.
            table.TableRows.Append(row);
            int rowCount = table.TableRows.Count;

            //Get the last row.
            TableRow lastRow = table.TableRows[rowCount - 1];

            //Set new data of the first cell of last row.
            lastRow[0].TextFrame.Text = " The first added cell";

            //Set new data of the second cell of last row.
            lastRow[1].TextFrame.Text = " The second added cell";

            String result = "Result-AddRowToTable.pptx";

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