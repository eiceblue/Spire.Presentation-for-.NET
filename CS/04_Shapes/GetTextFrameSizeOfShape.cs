using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetTextFrameSizeOfShape
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetTextFrameSizeOfShape.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            StringBuilder sb = new StringBuilder();

            // Iterate the shapes in the slide
            for (int i=0;i<slide.Shapes.Count;i++)
            {

                IAutoShape autoShape = slide.Shapes[i] as IAutoShape;
                SizeF size =  autoShape.TextFrame.GetTextSize();
                sb.AppendLine("The size of text frame in shape" + i + ", width:" + size.Width + " height:" +size.Height);
            }

            File.WriteAllText("GetTextFrameSizeOfShape.txt", sb.ToString());

            System.Diagnostics.Process.Start("GetTextFrameSizeOfShape.txt");

            presentation.Dispose();
        }
    }
}