using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace RemoveNoteAtSpecificSlide
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveNoteFromSlides.pptx");
            //Get the first slide
            ISlide slide = presentation.Slides[0];
            
            //Get note slide
            NotesSlide note = slide.NotesSlide;
            //Clear note text
            note.NotesTextFrame.Text = "";

            String result = "RemoveNotesAtSpecificSlide_result.pptx";
            //Save the PPT to PDF file format
            presentation.SaveToFile(result, FileFormat.Pptx2007);

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