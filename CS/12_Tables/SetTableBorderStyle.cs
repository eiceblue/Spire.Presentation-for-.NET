using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SetTableBorderStyle
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

            //Find the table by looping through all the slides, and then set borders for it. 
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is ITable)
                    {
                        foreach (TableRow row in (shape as ITable).TableRows)
                        {
                            foreach (Cell cell in row)
                            {
                                cell.BorderTop.FillType = FillFormatType.Solid;
                                cell.BorderBottom.FillType = FillFormatType.Solid;
                                cell.BorderLeft.FillType = FillFormatType.Solid;
                                cell.BorderRight.FillType = FillFormatType.Solid;
                            }
                        }
                    }
                }
            }

            String result = "Result-SetTableBorderStyle.pptx";

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