using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace RemoveTableFromPptSlide
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

            //Get the tables within the PPT document.
            List<IShape> shape_tems = new List<IShape>();

            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    //Add new table to table list.
                    shape_tems.Add(shape);
                }
            }

            //Remove all the tables form the first slide.
            foreach (IShape shape in shape_tems)
            {
                presentation.Slides[0].Shapes.Remove(shape);
            }

            String result = "Result-RemoveTableFromPptSlide.pptx";

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