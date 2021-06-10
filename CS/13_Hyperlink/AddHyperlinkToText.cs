using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddHyperlinkToText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document.
            Presentation presentation = new Presentation();

            //Load the file from disk.
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\AddHyperlinkToText.pptx");

            //Find the text we want to add link to it.
            IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;
            TextParagraph tp = shape.TextFrame.TextRange.Paragraph;
            string temp = tp.Text;

            //Split the original text.
            string textToLink = "Spire.Presentation";
            string[] strSplit = temp.Split(new string[] { "Spire.Presentation" }, StringSplitOptions.None);

            //Clear all text.
            tp.TextRanges.Clear();

            //Add new text.
            TextRange tr = new TextRange(strSplit[0]);
            tp.TextRanges.Append(tr);

            //Add the hyperlink.
            tr = new TextRange(textToLink);
            tr.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
            tp.TextRanges.Append(tr);

            String result = "Result-AddHyperlinkToText.pptx";

            //Save to file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the PowerPoint file.
            PptDocumentViewer(result);
        }

        private void PptDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}