using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace EditTableDataAndStyle
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

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_1.pptx");

            //Store the data used in replacement in string [].
            string[] str = new string[] { "Germany", "Berlin", "Europe", "0152458", "20860000" };

            ITable table = null;

            //Get the table in PowerPoint document.
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Change the style of table.
                    table.StylePreset = TableStylePreset.LightStyle1Accent2;

                    for (int i = 0; i < table.ColumnsList.Count; i++)
                    {
                        //Replace the data in cell.
                        table[i, 2].TextFrame.Text = str[i];

                        //Set the highlightcolor.
                        table[i, 2].TextFrame.TextRange.HighlightColor.Color = Color.BlueViolet;
                    }
                }
            }

            String result = "Result-EditTableDataAndStyle.pptx";

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