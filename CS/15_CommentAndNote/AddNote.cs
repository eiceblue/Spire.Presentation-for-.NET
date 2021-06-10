using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddNote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PPT document and load file
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\AddNote.pptx");
           
            ISlide slide = ppt.Slides[0];

            //Add note slide
            NotesSlide notesSlide = slide.AddNotesSlide();

            //Add paragraph in the notesSlide
            TextParagraph paragraph = new TextParagraph();
            paragraph.Text = "Tips for making effective presentations:";
            notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);

            paragraph = new TextParagraph();
            paragraph.Text = "Use the slide master feature to create a consistent and simple design template.";
            notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);
            //Set the bullet type for the paragraph in notesSlide
            notesSlide.NotesTextFrame.Paragraphs[1].BulletType = TextBulletType.Numbered;
            notesSlide.NotesTextFrame.Paragraphs[1].BulletStyle = NumberedBulletStyle.BulletArabicPeriod;

            paragraph = new TextParagraph();
            paragraph.Text = "Simplify and limit the number of words on each screen.";
            notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);
            notesSlide.NotesTextFrame.Paragraphs[2].BulletType = TextBulletType.Numbered;
            notesSlide.NotesTextFrame.Paragraphs[2].BulletStyle = NumberedBulletStyle.BulletArabicPeriod;

            paragraph = new TextParagraph();
            paragraph.Text = "Use contrasting colors for text and background.";
            notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);
            notesSlide.NotesTextFrame.Paragraphs[3].BulletType = TextBulletType.Numbered;
            notesSlide.NotesTextFrame.Paragraphs[3].BulletStyle = NumberedBulletStyle.BulletArabicPeriod;

            //Save the file
            ppt.SaveToFile("AddNote.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("AddNote.pptx");
        }

       
    }
}
