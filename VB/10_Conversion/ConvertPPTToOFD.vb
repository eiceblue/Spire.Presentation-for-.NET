Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace ConvertPPTToOFD
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Load ppt file
			presentation.LoadFromFile("..\..\..\..\..\..\Data\CopyParagraph.pptx")

			'Save the PPT document to OFD format
			Dim result As String = "ConvertPPTToOFD_result.ofd"
			presentation.SaveToFile(result, Spire.Presentation.FileFormat.OFD)

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