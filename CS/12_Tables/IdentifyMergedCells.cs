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

namespace IdentifyMergedCells
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

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\MergedCellInTable.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            StringBuilder str = new StringBuilder();
            String output="";
            foreach (IShape shape in slide.Shapes)
            {
                //Verify if it is table
                if (shape is ITable)
                {
                    ITable table = (ITable)shape;
                    for (int r = 0; r < table.TableRows.Count; r++)
                    {
                        for (int c = 0; c < table.ColumnsList.Count; c++)
                        {
                            // Get cell
                            Cell currentCell = table.TableRows[r][c];
                            //Identify if it is merged cell
                            if (currentCell.RowSpan>1 || currentCell.ColSpan>1)
                            {
                                output =string.Format("Cell {0}:{1} is a part of merged cell with RowSpan={2} and ColSpan={3} starting from Cell {4}:{5}.",
                                                  r, c, currentCell.RowSpan, currentCell.ColSpan, currentCell.FirstRowIndex, currentCell.FirstColumnIndex);

                                str.AppendLine(output);
                            }
                        }

                    }                  
                }
            }

            string result = "IdentifyMergedCells_result.txt";
            File.WriteAllText(result, str.ToString());
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