using System;
using System.Collections;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Collections;

namespace GetOLEProperties
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetOLEPropertiesOutsideOfShape.pptx");

            //Get the first slide
            ISlide slide = presentation.Slides[0];

            //Get the first OLE
            OleObjectCollection oles = slide.OleObjects;
            OleObject oleObject = oles[0];

            StringBuilder sb = new StringBuilder();

            //Get the information of OLE Object
            sb.AppendLine("ShapeID=" + oleObject.ShapeID);
            sb.AppendLine("FrameTop=" + oleObject.Frame.Top);
            sb.AppendLine("FrameLeft=" + oleObject.Frame.Left);
            sb.AppendLine("FrameWidth=" + oleObject.Frame.Width);
            sb.AppendLine("FrameHight=" + oleObject.Frame.Height);
            sb.AppendLine("IsHidden=" + oleObject.IsHidden);

            //Get the properties of OLE
            foreach (DictionaryEntry entry in oleObject.Properties)
            {
                sb.AppendLine(entry.Key + ":" + entry.Value);
            }

            // Save and preview the output file
            File.AppendAllText("GetOLEOutsideOfShape.txt", sb.ToString());

            System.Diagnostics.Process.Start("GetOLEOutsideOfShape.txt");

            presentation.Dispose();
        }
    }
}