using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace GetEmbeddedFileName
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

            //Load document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\oleTest.pptx");

            //Loop through the slides and shapes
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is IOleObject)
                    {
                        //Find OLE object
                        IOleObject oleObject = shape as IOleObject;
                        //Get OLE object label name
                        string oleFileName = oleObject.EmbeddedFileName;
                        MessageBox.Show("The name of the OLE object label is:"+oleFileName);
                    }
                }
            }
         
            
          
        }
    }
}
