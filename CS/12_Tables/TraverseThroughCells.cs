using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace TraverseThroughCells
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPonit document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_1.pptx");

            StringBuilder content = new StringBuilder();
            content.AppendLine("The data in cells of this PowerPoint file is: ");

            //Get the table.
            ITable table = null;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Traverse through the cells of table.
                    foreach (TableRow row in table.TableRows)
                    {
                        foreach (Cell cell in row)
                        {
                            content.AppendLine(cell.TextFrame.Text);
                        }
                        content.AppendLine("\n");
                    }
                }
            }

            String result = "Result-TraverseThroughCells.txt";

            //Save to file.
            File.WriteAllText(result, content.ToString());

            //Launch the file.
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