using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace DetectUsedThemes
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Themes.pptx");

            StringBuilder sb = new StringBuilder();
            string themeName = null;
            sb.AppendLine("This is the name list of the used theme below.");
            //Get the theme name of each slide in the document
            foreach (ISlide slide in ppt.Slides)
            {
                themeName = slide.Theme.Name;
                sb.AppendLine(themeName);
            }

            //Save to the text document
            string result = "DetectUsedThemes.txt";
            File.WriteAllText(result, sb.ToString());
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