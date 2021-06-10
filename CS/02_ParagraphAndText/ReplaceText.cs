using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Spire.Presentation;

namespace ReplaceText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> tagValues = new Dictionary<string, string>();
            tagValues.Add("Spire.Presentation for .NET", "Spire.PPT");

            //Create an instance of presentation document
            Presentation ppt = new Presentation();
            //Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\TextTemplate.pptx");

            ReplaceTags(ppt.Slides[0], tagValues);

            //Save the document
            string result = "ReplaceText.pptx";
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

        private void ReplaceTags(ISlide pSlide, Dictionary<string, string> TagValues)
        {
            foreach (IShape curShape in pSlide.Shapes)
            {
                if (curShape is IAutoShape)
                {
                    foreach (TextParagraph tp in (curShape as IAutoShape).TextFrame.Paragraphs)
                    {
                        foreach (var curKey in TagValues.Keys)
                        {
                            if (tp.Text.Contains(curKey))
                            {
                                tp.Text = tp.Text.Replace(curKey, TagValues[curKey]);
                            }
                        }
                    }
                }
            }
        }
    }
}