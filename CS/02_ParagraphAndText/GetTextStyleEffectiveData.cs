using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing.Transition;
using Spire.Presentation.Diagrams;
using System.IO;
using Spire.Presentation.Drawing;

namespace GetTextStyleEffectiveData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document
            Presentation presentation = new Presentation();

            //Load PPT file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Az1.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            //Get a shape 
            IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;
 
            StringBuilder str = new StringBuilder();
            for (int p = 0; p < shape.TextFrame.Paragraphs.Count; p++)
            {   
                var paragraph = shape.TextFrame.Paragraphs[p];
                str.AppendLine("Text style for Paragraph " + p + " :");
                //Get the paragraph style
                str.AppendLine(" Indent: " + paragraph.Indent);
                str.AppendLine(" Alignment: " + paragraph.Alignment);
                str.AppendLine(" Font alignment: " + paragraph.FontAlignment);
                str.AppendLine(" Hanging punctuation: " + paragraph.HangingPunctuation);
                str.AppendLine(" Line spacing: " + paragraph.LineSpacing);
                str.AppendLine(" Space before: " + paragraph.SpaceBefore);
                str.AppendLine(" Space after: " + paragraph.SpaceAfter.ToString());
                str.AppendLine();
                for (int r = 0; r < paragraph.TextRanges.Count; r++)
                {                
                    var textRange = paragraph.TextRanges[r];
                    str.AppendLine("  Text style for Paragraph " + p + " TextRange " + r + " :");
                    //Get the text range style
                    str.AppendLine("    Font height: " + textRange.FontHeight);
                    str.AppendLine("    Language: " + textRange.Language);
                    str.AppendLine("    Font: " + textRange.LatinFont.FontName);
                    str.AppendLine();
                }
            }

            string result = "GetTextStyleEffectiveData_result.txt";
            File.WriteAllText(result, str.ToString());
            Viewer(result);
        }

        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}