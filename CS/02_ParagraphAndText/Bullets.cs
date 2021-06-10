using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;


namespace Bullets
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Bullets.pptx");

            IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[1];

            foreach (TextParagraph para in shape.TextFrame.Paragraphs)
            {
                //Add the bullets
                para.BulletType = TextBulletType.Numbered;
                para.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod;

            }

            //Save the document and launch
            presentation.SaveToFile("bullets.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("bullets.pptx");
        }
    }
}