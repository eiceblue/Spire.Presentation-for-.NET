using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace SaveSvgWithOption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {            
            string inputFile = @"..\..\..\..\..\..\Data\SaveSvgWithOption.pptx";

            // Create Presentation object and load the file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(inputFile);

            // Save the underline as decoration when converting to Svg
            ppt.SaveToSvgOption.SaveUnderlineAsDecoration = true;

            // Save to Svg
            byte[] svgByte = ppt.Slides[0].Shapes[0].SaveAsSvgInSlide();
            FileStream fs = new FileStream("SaveSvgWithOption" + "1.svg", FileMode.Create);
            fs.Write(svgByte, 0, svgByte.Length);
            fs.Close();

            //Dispose
            ppt.Dispose();

        }

    }
}