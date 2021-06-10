using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace Properties
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Properties.pptx");

            //Set the DocumentProperty of PPT document
            presentation.DocumentProperty.Application = "Spire.Presentation";
            presentation.DocumentProperty.Author = "E-iceblue";
            presentation.DocumentProperty.Company = "E-iceblue Co., Ltd.";
            presentation.DocumentProperty.Keywords = "Demo File";
            presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
            presentation.DocumentProperty.Category = "Demo";
            presentation.DocumentProperty.Title = "This is a demo file.";
            presentation.DocumentProperty.Subject = "Test";

            //Save the document
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

            //Launch the PPT file
            System.Diagnostics.Process.Start("Output.pptx");
        }
    }
}