using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetAlignmentInTable
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SetAlignmentInTable.pptx");

            ITable table = null;
            foreach (IShape shape in presentation.Slides[0].Shapes)
            {
                if (shape is ITable)
                {
                    table = (ITable)shape;

                    //Horizontal Alignment
                    //Set the horizontal alignment for the cells in first column 
                    table[0, 1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
                    table[0, 2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                    table[0, 3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right;
                    table[0, 4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify;

                    //Vertical Alignment
                    //Set the vertical alignment for the cells in second column 
                    table[1, 1].TextAnchorType = TextAnchorType.Top;
                    table[1, 2].TextAnchorType = TextAnchorType.Center;
                    table[1, 3].TextAnchorType = TextAnchorType.Bottom;
                    table[1, 4].TextAnchorType = TextAnchorType.None;

                    //Both orientaions
                    //Set the both horizontal and vertical alignment for the cells in the third column 
                    table[2, 1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
                    table[2, 1].TextAnchorType = TextAnchorType.Top;

                    table[2, 2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right;
                    table[2, 2].TextAnchorType = TextAnchorType.Center;

                    table[2, 3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify;
                    table[2, 3].TextAnchorType = TextAnchorType.Bottom;

                    table[2, 4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
                    table[2, 4].TextAnchorType = TextAnchorType.Top;
                }
            }

            //Save the document
            presentation.SaveToFile("SetAlignmentInTable_result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("SetAlignmentInTable_result.pptx");
        }
    }
}
