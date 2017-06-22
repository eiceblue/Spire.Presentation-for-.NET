using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddImageInTableCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //create a PPT document
            Presentation presentation = new Presentation();       

            //create a table and set the table style
            Double[] widths = new double[] { 100, 100 };
            Double[] heights = new double[] { 100, 100 };
            ITable table = presentation.Slides[0].Shapes.AppendTable(100, 80, widths, heights);
            table.StylePreset = TableStylePreset.LightStyle1Accent2;

            //load the image and insert it into table
            IImageData imgData = presentation.Images.Append(Image.FromFile(@"..\..\..\..\..\..\Data\flower.jpg"));
            table[0, 0].FillFormat.FillType = FillFormatType.Picture;
            table[0, 0].FillFormat.PictureFill.Picture.EmbedImage = imgData;
            table[0, 0].FillFormat.PictureFill.FillType = PictureFillType.Stretch;

            //save the document
            presentation.SaveToFile("InsertImageInTable.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("InsertImageInTable.pptx");
        }
    }
}
