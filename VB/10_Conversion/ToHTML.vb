Imports Spire.Presentation

Namespace ToHTML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()

			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Conversion.pptx")

			'Save the document to HTML format
			Dim result As String = "ToHTML.html"
			ppt.SaveToFile(result, FileFormat.Html)
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