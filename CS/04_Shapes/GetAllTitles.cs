using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetAllTitles
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
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Titles.pptx");

            //Instantiate a list of IShape objects
            List<IShape> shapelist = new List<IShape>();
            //Loop through all sildes and all shapes on each slide
            foreach (ISlide slide in ppt.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape.Placeholder != null)
                    {
                        //Get all titles
                        switch (shape.Placeholder.Type)
                        {
                            case PlaceholderType.Title:
                                shapelist.Add(shape);
                                break;
                            case PlaceholderType.CenteredTitle:
                                shapelist.Add(shape);
                                break;
                            case PlaceholderType.Subtitle:
                                shapelist.Add(shape);
                                break;
                        }
                    }
                }
            }

            //Loop through the list and get the inner text of all shapes in the list
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Below are all the obtained titles:");
            for (int i = 0; i < shapelist.Count; i++)
            {
                IAutoShape shape1 = shapelist[i] as IAutoShape;
                sb.AppendLine(shape1.TextFrame.Text);
            }

            //Save to the Text file
            string result = "GetAllTitles.txt";
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