using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace AddAndGetSpeakerNotes
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_1.pptx");

            //Get the first slide and in the PowerPoint document.
            ISlide slide = presentation.Slides[0];

            //Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
            NotesSlide ns = slide.NotesSlide;
            if (ns == null)
            {
                ns = slide.AddNotesSlide();
            }

            //Add the text string as the notes.
            ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation";

            StringBuilder content = new StringBuilder();
            content.AppendLine("The speaker notes added by Spire.Presentation is: " + ns.NotesTextFrame.Text);

            String result = "Result-AddAndGetSpeakerNotes.pptx";
            String result1 = "Result-AddAndGetSpeakerNotes.txt";

            //Save to PowerPoint file.
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Get the speaker notes and save to txt file.
            File.WriteAllText(result1,content.ToString());

            //Launch the PowerPoint file.
            PptDocumentViewer(result);

            //Launch the txt file.
            PptDocumentViewer(result1);
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