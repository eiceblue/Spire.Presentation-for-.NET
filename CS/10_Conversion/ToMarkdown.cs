using System;
using System.Windows.Forms;
using Spire.Presentation;

namespace ToMarkdown

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create and load the file 
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractText.pptx");
            // Convert to markdown format
            ppt.SaveToFile("ToMarkdown.md", FileFormat.Markdown);
            ppt.Dispose();

        }
    }
}