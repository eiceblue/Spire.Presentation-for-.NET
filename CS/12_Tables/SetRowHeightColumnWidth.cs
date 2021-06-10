using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetRowHeightColumnWidth
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\SetRowHeightColumnWidth.pptx");

            //Get the table
            ITable table = null;
            foreach (IShape shape in ppt.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Set the height for the rows
                    table.TableRows[0].Height = 100;
                    table.TableRows[1].Height = 80;
                    table.TableRows[2].Height = 60;
                    table.TableRows[3].Height = 40;
                    table.TableRows[4].Height = 20;

                    //Set the column width
                    table.ColumnsList[0].Width = 60;
                    table.ColumnsList[1].Width = 80;
                    table.ColumnsList[2].Width = 120;
                    table.ColumnsList[3].Width = 140;
                    table.ColumnsList[4].Width = 160;
                }
            }
            //Save the file
            ppt.SaveToFile("RowHeightAndColumnWidth_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("RowHeightAndColumnWidth_result.pptx");
        }
    }
}
