using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using System.IO;
using Spire.Presentation;
using Spire.Presentation.Diagrams;
using Spire.Presentation.Charts;

namespace OperatePlaceholders
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

            //Load the document from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\OperatePlaceholders.pptx");

            //Operate placeholders
            for (int j=0;j<presentation.Slides.Count;j++)
            {
                ISlide slide = (ISlide)presentation.Slides[j];
                
                for (int i=0;i<slide.Shapes.Count;i++)
                {
                    Shape shape = (Shape)slide.Shapes[i];
                    switch(shape.Placeholder.Type)
                    {
                        case PlaceholderType.Media:
                            shape.InsertVideo(@"..\..\..\..\..\..\Data\Video.mp4");
                            break;
                       
                        case PlaceholderType.Picture:
                            shape.InsertPicture( @"..\..\..\..\..\..\Data\E-iceblueLogo.png");
                            break;
                        
                        case PlaceholderType.Chart:
                            shape.InsertChart(ChartType.ColumnClustered);
                            break;
                        
                        case PlaceholderType.Table:
                            shape.InsertTable(3,2);
                            break;
                        
                        case PlaceholderType.Diagram:
                            shape.InsertSmartArt(SmartArtLayoutType.BasicBlockList);
                            break;
                    }
                }
            }
 
            string result="OperatePlaceholders_result.pptx";
            //Save the document
            presentation.SaveToFile(result, FileFormat.Pptx2013);

            //Launch the file
            PPTDocumentViewer(result);
        }
        private void PPTDocumentViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }

        }
    }
}