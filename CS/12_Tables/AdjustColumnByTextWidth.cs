using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace AdjustColumnByTextWidth
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

            //Get the table from the first slide of the sample document.
            ISlide slide = presentation.Slides[0];
            ITable table = slide.Shapes[0] as ITable;

            //Adjust the first column width of table by text width.
            table.ColumnsList[0].AdjustColumnByTextWidth();



            //Save to file.
            String result = "output.pptx";
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