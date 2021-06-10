using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace RemoveTableBorderStyle
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
                                cell.BorderTop.FillType = FillFormatType.None;
                                cell.BorderBottom.FillType = FillFormatType.None;
                                cell.BorderLeft.FillType = FillFormatType.None;
                                cell.BorderRight.FillType = FillFormatType.None;
                            }
                        }
                    }
                }
            }

            String result = "Result-RemoveTableBorderStyle.pptx";

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