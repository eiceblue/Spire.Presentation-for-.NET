using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace EmbedZipIntoPPT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a ppt document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\EmbedZipIntoPPT.pptx");

            //Load a zip object
            String zipPath = @"..\..\..\..\..\..\Data\test.zip";
            byte[] data = File.ReadAllBytes(zipPath);

            Rectangle rec = new Rectangle(80, 60, 100, 100);

            //Insert the zip object to presentation
            IOleObject ole = ppt.Slides[0].Shapes.AppendOleObject(@"test.zip", data, rec);
            ole.ProgId = "Package";
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\icon.png");
            IImageData oleImage = ppt.Images.Append(image);
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            FileStream stream = new FileStream(@"..\..\..\..\..\..\Data\icon.png", FileMode.Open);
            IImageData oleImage = ppt.Images.Append(stream);
            */
            
            ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;

            //Save the document
            ppt.SaveToFile("EmbedZipIntoPPT_result.pptx", FileFormat.Pptx2010);

            //Launch the file
            OutputViewer("EmbedZipIntoPPT_result.pptx");
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