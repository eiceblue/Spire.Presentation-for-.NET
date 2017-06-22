using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace EmbedExcelAsOLE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\EmbedExcelAsOLE.png");
            Presentation ppt = new Presentation();
            IImageData oleImage = ppt.Images.Append(image);
            Rectangle rec = new Rectangle(60, 60, image.Width, image.Height);

            //insert an OLE object to presentation based on the Excel data

            Spire.Presentation.IOleObject oleObject = ppt.Slides[0].Shapes.AppendOleObject("excel", File.ReadAllBytes(@"..\..\..\..\..\..\Data\DatatableSample.xlsx"), rec);
            oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
            oleObject.ProgId = "Excel.Sheet.12";


            //save the document
            ppt.SaveToFile("InsertOle.pptx", Spire.Presentation.FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("InsertOle.pptx");
        }
    }
}
