using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace SetColumnsCountOfTextFrame
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
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ColumnsCount.pptx");

            //Get the first shape in first slide and set column count of text for it.
            IAutoShape shape1 = (IAutoShape)ppt.Slides[0].Shapes[0];
            shape1.TextFrame.ColumnCount = 3;

            //Get the second shape in second slide and set column count of text for it.
            IAutoShape shape2 = (IAutoShape)ppt.Slides[1].Shapes[0];
            shape2.TextFrame.ColumnCount = 2;

            //Save the document
            ppt.SaveToFile("result.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("result.pptx");
        }
    }
}


