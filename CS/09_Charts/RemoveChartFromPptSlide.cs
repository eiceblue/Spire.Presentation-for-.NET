using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace RemoveChartFromPptSlide
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPonit document
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_3.pptx");

            //Get the first slide from the document.
            ISlide slide = presentation.Slides[0];

            //Remove chart from the slide.
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                IShape shape = slide.Shapes[i] as IShape;
                if (shape is IChart)
                {
                    slide.Shapes.Remove(shape);
                }
            }

            String result = "Result-RemoveChartFromPptSlide.pptx";

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