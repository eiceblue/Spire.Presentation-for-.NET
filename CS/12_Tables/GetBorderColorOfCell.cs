using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;

namespace GetBorderColorOfCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a PowerPoint document
            Presentation presentation = new Presentation();

            //Load file from disk
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetBorderColorOfCell.pptx");

            //Get the table in the first slide
            ITable table = presentation.Slides[0].Shapes[0] as ITable;

            //Get borders' color of the first cell
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Color of left border:" + table[0, 0].BorderLeftDisplayColor);
            sb.AppendLine("Color of top border:" + table[0, 0].BorderTopDisplayColor);
            sb.AppendLine("Color of right border:" + table[0, 0].BorderRightDisplayColor);
            sb.AppendLine("Color of bottom border:" + table[0, 0].BorderBottomDisplayColor);

            //Get display color of the first cell
            sb.AppendLine("Color of cell:"+ table[0,0].DisplayColor);
            string result = "Result-SetChartDataLabelRange.txt";

            File.WriteAllText(result, sb.ToString()) ;
            System.Diagnostics.Process.Start(result);
        }


    }
}