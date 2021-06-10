using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace InsertEMFInPPT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\BlankSample_N.pptx");

            //EMF file path
            string ImageFile = @"..\..\..\..\..\..\Data\InsertEMF.emf";

            //Define image size
            Image img = Image.FromFile(ImageFile);
            float width=img.Width/1.5f;
            float height=img.Height/1.5f;
            RectangleF rect = new RectangleF(100, 100, width,height);

            //Append the EMF in slide
            IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
            image.Line.FillType = FillFormatType.None;

            //Save the document
            string result = "InsertEMFInPPT_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}