using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace MixFontStyles
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
            //Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\FontStyle.pptx");

            //Get the second shape of the first slide
            IAutoShape shape = ppt.Slides[0].Shapes[1] as IAutoShape;
            //Get the text from the shape 
            string originalText = shape.TextFrame.Text;

            //Split the string by specified words and return substrings to a string array
            string[] splitArray = originalText.Split(new string[] { "bold", "red", "underlined", "bigger font size" }, StringSplitOptions.None);

            //Remove the paragraph from TextRange
            TextParagraph tp = shape.TextFrame.TextRange.Paragraph;
            tp.TextRanges.Clear();

            //Append normal text that is in front of 'bold' to the paragraph
            TextRange tr = new TextRange(splitArray[0]);
            tp.TextRanges.Append(tr);
            //Set font style of the text 'bold' as bold
            tr = new TextRange("bold");
            tr.IsBold = TriState.True;
            tp.TextRanges.Append(tr);

            //Append normal text that is in front of 'red' to the paragraph
            tr = new TextRange(splitArray[1]);
            tp.TextRanges.Append(tr);
            //Set the color of the text 'red' as red
            tr = new TextRange("red");
            tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            tr.Format.Fill.SolidColor.Color = Color.Red;
            tp.TextRanges.Append(tr);

            //Append normal text that is in front of 'underlined' to the paragraph
            tr = new TextRange(splitArray[2]);
            tp.TextRanges.Append(tr);
            //Underline the text 'undelined'
            tr = new TextRange("underlined");
            tr.TextUnderlineType = TextUnderlineType.Single;
            tp.TextRanges.Append(tr);

            //Append normal text that is in front of 'bigger font size' to the paragraph
            tr = new TextRange(splitArray[3]);
            tp.TextRanges.Append(tr);
            //Set a large font for the text 'bigger font size'
            tr = new TextRange("bigger font size");
            tr.FontHeight = 35;
            tp.TextRanges.Append(tr);

            //Append other normal text
            tr = new TextRange(splitArray[4]);
            tp.TextRanges.Append(tr);

            //Save the document
            string result = "MixFontStyles.pptx";
            ppt.SaveToFile(result, FileFormat.Pptx2013);
            PresentationDocViewer(result);
		}
	
		private void PresentationDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}