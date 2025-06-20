using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Presentation;

namespace InsertPlaceholderInMaster
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {            
            // Create a Presentation object
            Presentation presentation = new Presentation();

             // Inset palce hodler 
            presentation.Masters[0].Layouts[0].InsertPlaceholder(InsertPlaceholderType.Text, new RectangleF(20, 30, 400, 400));
            
            // Save file 
            presentation.SaveToFile("InsertPlaceholderInMaster_output.pptx", FileFormat.Pptx2019);

            //Dispose
            presentation.Dispose();

            System.Diagnostics.Process.Start("InsertPlaceholderInMaster_output.pptx");
        }
    }
}