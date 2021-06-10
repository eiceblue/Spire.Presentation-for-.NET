using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace RemoveTextBox
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\TextBoxTemplate.pptx");

            //Get the first slide
            ISlide slide = ppt.Slides[0];
            //Traverse all the shapes in slide
            for (int i = 0; i < slide.Shapes.Count;i++)
            {
                if(slide.Shapes[i].Name.Contains("TextBox"))
                {
                	slide.Shapes.RemoveAt(i);
                	i--;
                }
            }

            //Save the document
            string result = "RemoveTextBox.pptx";
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