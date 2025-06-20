using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;

namespace ModifyStyleOfFirstFoundText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create an instance of presentation document
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\TextTemplate.pptx");

            //Find first "Spire"
            string text = "Spire";
            TextRange textRange = ppt.Slides[0].FindFirstTextAsRange(text);
            
            //Modify the style
            textRange.Fill.FillType = FillFormatType.Solid;
            textRange.Fill.SolidColor.Color = Color.Red;
            textRange.FontHeight = 28;
            textRange.LatinFont = new TextFont("Calibri");
            textRange.IsBold = TriState.True;
            textRange.IsItalic = TriState.True;
            textRange.TextUnderlineType = TextUnderlineType.Double;
            textRange.TextStrikethroughType = TextStrikethroughType.Single;

            //Save the document
            string result = "Result.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
}
	
	private static void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}