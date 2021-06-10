using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace CopyChartBetweenPptFiles
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
            Presentation presentation1 = new Presentation();

            //Load the file from disk.
            presentation1.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_2.pptx");

            //Get the chart that is going to be copied.
            IChart chart = presentation1.Slides[0].Shapes[0] as IChart;

            //Load the second PowerPoint document.
            Presentation presentation2 = new Presentation();
            presentation2.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_1.pptx");

            //Copy chart from the first document to the second document.
            presentation2.Slides.Append();
            presentation2.Slides[1].Shapes.CreateChart(chart, new RectangleF(100, 100, 500, 300), -1);

            String result = "Result-CopyChartBetweenPptFiles.pptx";

            //Save to file.
            presentation2.SaveToFile(result, FileFormat.Pptx2013);

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