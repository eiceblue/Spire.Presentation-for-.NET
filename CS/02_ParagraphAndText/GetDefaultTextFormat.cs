using System;
using System.IO;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace GetDefaultTextFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {            
            string inputFile = @"..\..\..\..\..\..\Data\GetDefaultTextFormat.pptx";
            string outputFile = "GetDefaultTextFormat.txt";

            // Create Presentation object and load the file
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(inputFile);

            // Get the first shape of the first slide
            IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

            // Get the display format of the text in shape
            DefaultTextRangeProperties format = shape.TextFrame.Paragraphs[0].TextRanges[0].DisPlayFormat;

            // Determine whether the format is bold or italic
            File.AppendAllText(outputFile, "Is the first shape text bolded :" + format.IsBold + "\r\n");
            File.AppendAllText(outputFile, "Is the first shape text italicized :" + format.IsItalic + "\r\n");

            // Dispose
            presentation.Dispose();

            System.Diagnostics.Process.Start(outputFile);

        }

    }
}