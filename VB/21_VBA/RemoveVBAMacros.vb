Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace RemoveVBAMacros
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Macros.ppt")
			'Remove macros
			'Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
			presentation.DeleteMacros()
			Dim result As String = "RemoveVBAMacros_result.ppt"
			presentation.SaveToFile(result,FileFormat.PPT)
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