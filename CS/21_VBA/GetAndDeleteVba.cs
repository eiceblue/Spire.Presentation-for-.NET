using System;
using System.Text;
using System.Windows.Forms;
using Spire.Presentation;
using Spire.Presentation.Converter.Vba;

namespace GetAndDeleteVba
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a Presentation object and load the PPTM file that contains VBA modules
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\VbaModule.pptm");

            //Get the VBA project of the loaded PPTM file
            IVbaProject vbaProject = ppt.VbaProject;

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("VBA project name: " + vbaProject.Name);
            stringBuilder.AppendLine("Total modules before deletion: " + vbaProject.Modules.Count);
            stringBuilder.AppendLine();

            //Read the VBA modules
            stringBuilder.AppendLine("[Modules Before Deletion]");
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                IVbaModule m = vbaProject.Modules[i];
                stringBuilder.AppendLine("  [" + i + "] Name: " + m.Name + ", Type: " + m.Type);
            }

            //Delete one module by its name (e.g. "ModuleB")
            vbaProject.Modules.Remove("ModuleB");
            stringBuilder.AppendLine();
            stringBuilder.AppendLine("Removed the module whose name is \"ModuleB\".");
            stringBuilder.AppendLine("Total modules after removing by name: " + vbaProject.Modules.Count);

            //Delete one module by its index (e.g. index 0)
            vbaProject.Modules.RemoveAt(0);
            stringBuilder.AppendLine();
            stringBuilder.AppendLine("Removed the module at index 0.");
            stringBuilder.AppendLine("Total modules after removing by index: " + vbaProject.Modules.Count);

            //Read the remaining VBA modules
            stringBuilder.AppendLine();
            stringBuilder.AppendLine("[Modules After Deletion]");
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                IVbaModule m = vbaProject.Modules[i];
                stringBuilder.AppendLine("  [" + i + "] Name: " + m.Name + ", Type: " + m.Type);
            }

            //Save the result as a single PPTM file
            ppt.SaveToFile("result.pptm", FileFormat.Pptm);

            //Write the operation log into a text file
            System.IO.File.WriteAllText("GetAndDeleteVba_out.txt", stringBuilder.ToString());
            ppt.Dispose();

            //Launch the log file so you can view the operation result
            System.Diagnostics.Process.Start("GetAndDeleteVba_out.txt");
        }
    }
}
