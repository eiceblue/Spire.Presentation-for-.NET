using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddImageInMaster
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

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddImageInMaster.pptx");

            //Get the master collection
            IMasterSlide master = presentation.Masters[0];

            //Append image to slide master
            String image = @"..\..\..\..\..\..\Data\Logo.png";
            RectangleF rff = new RectangleF(40, 40, 90, 90);
            IEmbedImage pic = master.Shapes.AppendEmbedImage(ShapeType.Rectangle, image, rff);
            pic.Line.FillFormat.FillType = FillFormatType.None;

            //Add new slide to presentation
            presentation.Slides.Append();

            //Save the document
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Output.pptx");
        }
    }
}
