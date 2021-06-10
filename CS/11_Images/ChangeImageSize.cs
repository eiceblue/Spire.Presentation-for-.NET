using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ChangeImageSize
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

            //Load document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractImage.pptx");

            //
            float scale=0.5f;
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is IEmbedImage)
                    {
                        IEmbedImage image = shape as IEmbedImage;
                        image.Width = image.Width * scale;
                        image.Height = image.Height * scale;
                    }
                }
            }
            
            //Save the document
            string result="ChangeImageSize_result.pptx";
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