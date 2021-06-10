Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace LoadFromStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Load PowerPoint file from stream
			Dim from_stream As FileStream = File.OpenRead("..\..\..\..\..\..\Data\InputTemplate.pptx")
			ppt.LoadFromStream(from_stream, FileFormat.Pptx2013)

			'Save the document
			Dim result As String = "LoadFromStream.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
			from_stream.Dispose()
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace