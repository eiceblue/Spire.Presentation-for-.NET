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
            //Load a PPT document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\AddImageInTableCell.pptx");

            //Get the first shape
            ITable table = ppt.Slides[0].Shapes[0] as ITable;

            //Load the image and insert it into table cell
            IImageData pptImg = ppt.Images.Append(Image.FromFile(@"..\..\..\..\..\..\Data\PresentationIcon.png"));

	//////////////////Use the following code for netstandard dlls/////////////////////////
            /*
			FileStream fileStream = new FileStream(@"..\..\..\..\..\..\Data\PresentationIcon.png", FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = new byte[fileStream.Length];
            fileStream.Read(bytes, 0, bytes.Length);
            fileStream.Close();
            Stream stream = new MemoryStream(bytes);
            IImageData pptImg = ppt.Images.Append(stream);
            stream.Close();
            fileStream.Close();
            */
            
            table[1, 1].FillFormat.FillType = FillFormatType.Picture;
            table[1, 1].FillFormat.PictureFill.Picture.EmbedImage = pptImg;
            table[1, 1].FillFormat.PictureFill.FillType = PictureFillType.Stretch;

            //Save the document
            ppt.SaveToFile("AddImageInTableCell_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddImageInTableCell_result.pptx");
        }
    }
}
