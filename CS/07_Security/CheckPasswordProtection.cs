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
using Spire.Presentation.Charts;

namespace CheckPasswordProtection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create Presentation
            Presentation presentation = new Presentation();

            //Check whether a PPT document is password protected
            bool isProtected=presentation.IsPasswordProtected(@"..\..\..\..\..\..\Data\Template_Ppt_4.pptx");
            
            //Show the result by message box
            MessageBox.Show("The file is " + (isProtected ? "password " : "not password ") +"protected!");
        }
    }
}