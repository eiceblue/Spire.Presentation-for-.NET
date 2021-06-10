using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace ExtractOLEObject
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractOLEObject.pptx");

            //Loop through the slides and shapes
            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is IOleObject)
                    {
                        //Find OLE object
                        IOleObject oleObject = shape as IOleObject;

                        //Get its data and write to file
                        byte[] bytes = oleObject.Data;
                        switch (oleObject.ProgId)
                        {
                            case "Excel.Sheet.8":
                                File.WriteAllBytes("result.xls", bytes);
                                break;
                            case "Excel.Sheet.12":
                                File.WriteAllBytes("result.xlsx", bytes);
                                break;
                            case "Word.Document.8":
                                File.WriteAllBytes("result.doc", bytes);
                                break;
                            case "Word.Document.12":
                                File.WriteAllBytes("result.docx", bytes);
                                break;
                            case "PowerPoint.Show.8":
                                File.WriteAllBytes("result.ppt", bytes);
                                break;
                            case "PowerPoint.Show.12":
                                File.WriteAllBytes("result.pptx", bytes);
                                break;
                        }
                    }
                }
            }
            MessageBox.Show("Completed!");
        }
    }
}