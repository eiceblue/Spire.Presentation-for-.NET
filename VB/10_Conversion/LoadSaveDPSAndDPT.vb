Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO

Namespace LoadSaveDPSAndDPT
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load .dps or .dpt file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\sample.dps",Spire.Presentation.FileFormat.Dps)
			 'presentation.LoadFromFile(@"..\..\..\..\..\..\Data\sample.dpt",Spire.Presentation.FileFormat.Dpt);

			'Save the .dps or .dpt file
			Dim result As String = "LoadSaveDPSAndDPT_result.dps"
			presentation.SaveToFile(result, Spire.Presentation.FileFormat.Dps)
			'presentation.SaveToFile("LoadSaveDPSAndDPT_result.dpt", Spire.Presentation.FileFormat.Dpt);
			'Launching the result file.
			Viewer(result)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace