using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace AddSlideToSection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string savePath = "AddSlideToSection.pptx";
            string input = @"..\..\..\..\..\..\Data\Section.pptx";

            //Create a PPT document
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(input);

            //Add a new shape to the PPT document
            presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(200, 50, 300, 100));

            //Create a new section and copy the first slide to it
            Section NewSection = presentation.SectionList.Append("New Section");
            NewSection.Insert(0, presentation.Slides[0]);

            presentation.SaveToFile(savePath, FileFormat.Pptx2013);
            System.Diagnostics.Process.Start(savePath);

        }
    }
}