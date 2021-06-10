using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace ManageNoteMasterHeaderFooter
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
            string loadPath = @"..\..\..\..\..\..\Data\PPTHasHeader.pptx";
            string savePath = "ManageNoteMasterHeaderFooter.pptx";

            //Load presentation
            presentation.LoadFromFile(loadPath);

            //Set the note Masters header and footer
            INoteMasterSlide noteMasterSlide = presentation.NotesMaster;
            if (!noteMasterSlide.Equals(null))
            {
                foreach(Shape shape in noteMasterSlide.Shapes)
                {
                    if (!shape.Placeholder.Equals(null))
                    {
                        if (shape.Placeholder.Type.Equals(PlaceholderType.Header))
                        {
                            (shape as IAutoShape).TextFrame.Text = "change the header by Spire";
                        }
                        if (shape.Placeholder.Type.Equals(PlaceholderType.Footer))
                        {
                            (shape as IAutoShape).TextFrame.Text = "change the footer by Spire";
                        }
                    }
                }
            }

            presentation.SaveToFile(savePath, FileFormat.Pptx2013);
            System.Diagnostics.Process.Start(savePath);
        }
    }
}