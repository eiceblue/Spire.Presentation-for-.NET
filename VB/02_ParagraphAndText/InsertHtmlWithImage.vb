Imports Spire.Presentation
Imports Spire.Presentation.Collections

Namespace InsertHtmlWithImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			Dim shapes As ShapeList = ppt.Slides(0).Shapes

			shapes.AddFromHtml("<html><div><p>First paragraph</p><p><img src='..\..\..\..\..\..\Data\Logo.png'/></p><p>Second paragraph </p></html>")

			'Save the document
			Dim result As String = "InsertHtmlWithImage.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
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