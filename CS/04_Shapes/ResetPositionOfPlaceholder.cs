using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ResetPositionOfPlaceholder
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_7.pptx");

            //Get the first slide from the sample document.
            ISlide slide = presentation.Slides[0];

            foreach (IShape shapeToMove in slide.Shapes)
            {
                //Reset the position of the slide number to the left.
                if (shapeToMove.Name.Contains("Slide Number Placeholder"))
                {
                    shapeToMove.Left = 0;
                }

                else if (shapeToMove.Name.Contains("Date Placeholder"))
                {
                    //Reset the position of the date time to the center.
                    shapeToMove.Left = presentation.SlideSize.Size.Width / 2;

                    //Reset the date time display style.
                    (shapeToMove as IAutoShape).TextFrame.TextRange.Paragraph.Text = DateTime.Now.ToString("dd.MM.yyyy");
                    (shapeToMove as IAutoShape).TextFrame.IsCentered = true;
                }
            }

            String result = "Result-ResetPositionOfDateTimeAndSlideNumber.pptx";

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