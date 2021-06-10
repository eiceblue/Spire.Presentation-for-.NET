using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using System.IO;

namespace ModifyOLEData
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ModifyOLEData.pptx");

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
                        MemoryStream pptStream = new MemoryStream(bytes);
                        MemoryStream stream=new MemoryStream();
                        if (oleObject.ProgId == "PowerPoint.Show.12")
                        {
                            //Load the PPT stream
                            Presentation ppt = new Presentation();
                            ppt.LoadFromStream(pptStream, Spire.Presentation.FileFormat.Auto);
                            //Append an image in slide
                            ppt.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, @"..\..\..\..\..\..\Data\Logo.png", new RectangleF(50, 50, 100, 100));
                            ppt.SaveToFile(stream, Spire.Presentation.FileFormat.Pptx2013);
                            stream.Position = 0;
                            //Modify the data
                            oleObject.Data = stream.ToArray();
                        }
                    }
                }
            }

            //Save the document
            string result = "ModifyOLEData_result.pptx";
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            OutputViewer(result);
        }
        private void OutputViewer(string filename)
        {
            try
            {
                System.Diagnostics.Process.Start(filename);
            }
            catch { }
        }
    }
}