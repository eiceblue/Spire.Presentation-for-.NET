using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace EditPromptText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string loadPath = @"..\..\..\..\..\..\Data\HasPromptText.pptx";
            string savePath = @"EditPromptText.pptx";
            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(loadPath);

            // Iterate through the slide
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape.Placeholder != null && shape is IAutoShape)
                {
                    string text = "";
                    // Set the text of the title
                    if (shape.Placeholder.Type == PlaceholderType.CenteredTitle)
                    {
                        text = "custom title create by Spire";
                    }
                    // Set text of the subtitle.
                    else if (shape.Placeholder.Type == PlaceholderType.Subtitle)
                    {
                        text = "custom subtitle create by Spire";
                    }

                    (shape as IAutoShape).TextFrame.Text = text;
                }
            }

            //Save the file
            presentation.SaveToFile(savePath, FileFormat.Pptx2013);
            System.Diagnostics.Process.Start(savePath);
        }
    }
}