using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation.Drawing;
using Spire.Presentation;

namespace CreateTable
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
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\CreateTable.pptx");
           
            Double[] widths = new double[] { 100, 100, 150, 100, 100 };
            Double[] heights = new double[] { 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15 };
            
            //Add new table to PPT
            ITable table = presentation.Slides[0].Shapes.AppendTable(presentation.SlideSize.Size.Width / 2 - 275, 90, widths, heights);

            String[,] dataStr = new String[,]{
            {"Name",	"Capital",	"Continent",	"Area",	"Population"},
            {"Venezuela",	"Caracas",	"South America",	"912047",	"19700000"},
            {"Bolivia",	"La Paz",	"South America",	"1098575",	"7300000"},
            {"Brazil",	"Brasilia",	"South America",	"8511196",	"150400000"},
            {"Canada",	"Ottawa",	"North America",	"9976147",	"26500000"},
            {"Chile",	"Santiago",	"South America",	"756943",	"13200000"},
            {"Colombia",	"Bagota",	"South America",	"1138907",	"33000000"},
            {"Cuba",	"Havana",	"North America",	"114524",	"10600000"},
            {"Ecuador",	"Quito",	"South America",	"455502",	"10600000"},
            {"Paraguay",	"Asuncion","South America", "406576",	"4660000"},
            {"Peru",	"Lima",	"South America",	"1285215",	"21600000"},
            {"Jamaica",	"Kingston",	"North America",	"11424",	"2500000"},
            {"Mexico",	"Mexico City",	"North America",	"1967180",	"88600000"}
            };

            //Add data to table
            for (int i = 0; i < 13; i++)
                for (int j = 0; j < 5; j++)
                {
                    //Fill the table with data
                    table[j, i].TextFrame.Text = dataStr[i, j];

                    //Set the Font
                    table[j, i].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Narrow");
                }

            //Set the alignment of the first row to Center
            for (int i = 0; i < 5; i++)
            {
                table[i, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
            }

            //Set the style of table
            table.StylePreset = TableStylePreset.LightStyle3Accent1;

            //Save the document
            presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);
            System.Diagnostics.Process.Start("Output.pptx");

        }
    }
}