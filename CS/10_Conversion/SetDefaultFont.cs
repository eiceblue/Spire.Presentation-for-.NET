using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetDefaultFont
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Specify the input and output files
            String inputFile = @"..\..\..\..\..\..\Data\ToPDF_DefaultFont.pptx";
            String outputFile = "SetDefaultFont.pdf";

            //Set the default fonts used for the conversion
            Presentation.SetDefaultLatinFontName("Arial");
            Presentation.SetDefaultEastAsianFontName("微软雅黑");

            //Convert the PPT document to PDF
            Presentation ppt=new Presentation();
            ppt.LoadFromFile(inputFile);
            ppt.SaveToFile(outputFile, FileFormat.PDF);

            //Reset the default font settings
            Presentation.ResetDefaultEastAsianFontName();
            Presentation.ResetDefaultLatinFontName();

            //Launch the PDF file
            System.Diagnostics.Process.Start(outputFile);
        }
    }
}
