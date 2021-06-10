Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace DeleteComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\DeleteComment.pptx")

			'Replace the text in the comment
			presentation.Slides(0).Comments(1).Text = "Replace comment"

			'Delete the third comment
			presentation.Slides(0).DeleteComment(presentation.Slides(0).Comments(2))

			'Save the document
			presentation.SaveToFile("DeleteComment.pptx", FileFormat.Pptx2010)
			Process.Start("DeleteComment.pptx")
		End Sub
	End Class
End Namespace
