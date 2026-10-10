Imports System
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation
Imports Spire.Presentation.Converter.Vba

Namespace GetAndDeleteVba
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create a Presentation object and load the PPTM file that contains VBA modules
            Dim ppt As Presentation = New Presentation()
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\VbaModule.pptm")

            'Get the VBA project of the loaded PPTM file
            Dim vbaProject As IVbaProject = ppt.VbaProject

            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendLine("VBA project name: " + vbaProject.Name)
            stringBuilder.AppendLine("Total modules before deletion: " + vbaProject.Modules.Count)
            stringBuilder.AppendLine()

            'Read the VBA modules
            stringBuilder.AppendLine("[Modules Before Deletion]")
            Dim i As Integer = 0
            While i < vbaProject.Modules.Count
                Dim m As IVbaModule = vbaProject.Modules[i]
                stringBuilder.AppendLine("  [" + i + "] Name: " + m.Name + ", Type: " + m.Type)
            Next

            'Delete one module by its name (e.g. "ModuleB")
            vbaProject.Modules.Remove("ModuleB")
            stringBuilder.AppendLine()
            stringBuilder.AppendLine("Removed the module whose name Is \"ModuleB\".")
            stringBuilder.AppendLine("Total modules after removing by name: " + vbaProject.Modules.Count)

            'Delete one module by its index (e.g. index 0)
            vbaProject.Modules.RemoveAt(0)
            stringBuilder.AppendLine()
            stringBuilder.AppendLine("Removed the module at index 0.")
            stringBuilder.AppendLine("Total modules after removing by index: " + vbaProject.Modules.Count)

            'Read the remaining VBA modules
            stringBuilder.AppendLine()
            stringBuilder.AppendLine("[Modules After Deletion]")
            Dim i As Integer = 0
            While i < vbaProject.Modules.Count
                Dim m As IVbaModule = vbaProject.Modules[i]
                stringBuilder.AppendLine("  [" + i + "] Name: " + m.Name + ", Type: " + m.Type)
            Next

            'Save the result as a single PPTM file
            ppt.SaveToFile("result.pptm", FileFormat.Pptm)

            'Write the operation log into a text file
            System.IO.File.WriteAllText("GetAndDeleteVba_out.txt", stringBuilder.ToString())
            ppt.Dispose()

            'Launch the log file so you can view the operation result
            System.Diagnostics.Process.Start("GetAndDeleteVba_out.txt")
        End Sub
    End Class
End Namespace
