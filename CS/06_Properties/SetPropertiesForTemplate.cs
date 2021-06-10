using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.IO;

namespace SetPropertiesForTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            //String for .pptx file 
            String pptxResult = "Output.pptx";

            //String for .odp file 
            String odpResult = "Output.odp";

            //String for .ppt file 
            String pptResult = "Output.ppt";

            //Create the .pptx template
            SetPropertiesForTemplate(pptxResult, FileFormat.Pptx2013);

            //Create the .odp template
            SetPropertiesForTemplate(odpResult, FileFormat.ODP);

            //Create the .ppt template
            SetPropertiesForTemplate(pptResult, FileFormat.PPT);

            //Launching the .pptx file.
            Viewer(pptxResult);
        }
        private static void SetPropertiesForTemplate(string filePath, FileFormat fileFormat)
        {
            //Create a document
            Presentation presentation = new Presentation();

            //Set the DocumentProperty 
            presentation.DocumentProperty.Application = "Spire.Presentation";
            presentation.DocumentProperty.Author = "E-iceblue";
            presentation.DocumentProperty.Company = "E-iceblue Co., Ltd.";
            presentation.DocumentProperty.Keywords = "Demo File";
            presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
            presentation.DocumentProperty.Category = "Demo";
            presentation.DocumentProperty.Title = "This is a demo file.";
            presentation.DocumentProperty.Subject = "Test";

            //Save to template file
            presentation.SaveToFile(filePath, fileFormat);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}