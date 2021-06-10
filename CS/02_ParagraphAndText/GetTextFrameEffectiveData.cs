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

namespace GetTextFrameEffectiveData
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

            ITextFrameProperties textFrameFormat = shape.TextFrame;
            StringBuilder str = new StringBuilder();
            str.AppendLine("Anchoring type: " + textFrameFormat.AnchoringType);
            str.AppendLine("Autofit type: " + textFrameFormat.AutofitType);
            str.AppendLine("Text vertical type: " + textFrameFormat.VerticalTextType);
            str.AppendLine("Margins");
            str.AppendLine("   Left: " + textFrameFormat.MarginLeft);
            str.AppendLine("   Top: " + textFrameFormat.MarginTop);
            str.AppendLine("   Right: " + textFrameFormat.MarginRight);
            str.AppendLine("   Bottom: " + textFrameFormat.MarginBottom);

            string result = "GetTextFrameEffectiveData_result.txt";
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