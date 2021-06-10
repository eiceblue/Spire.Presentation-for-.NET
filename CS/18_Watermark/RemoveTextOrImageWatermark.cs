using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace RemoveTextOrImageWatermark
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveTextAndImageWatermarks.pptx");

            //Remove text watermark by removing the shape which contains the text string "E-iceblue".
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                for (int j = 0; j < presentation.Slides[i].Shapes.Count; j++)
                {
                    if (presentation.Slides[i].Shapes[j] is IAutoShape)
                    {
                        IAutoShape shape = presentation.Slides[i].Shapes[j] as IAutoShape;
                        if (shape.TextFrame.Text.Contains("E-iceblue"))
                        {
                            presentation.Slides[i].Shapes.Remove(shape);
                        }
                    }
                }
            }

            //Remove image watermark.
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                presentation.Slides[i].SlideBackground.Fill.FillType = FillFormatType.None;
            }

            String result = "Result-RemoveTextAndImageWatermarks.pptx";

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