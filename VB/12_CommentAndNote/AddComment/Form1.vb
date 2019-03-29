Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace AddComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddComment.pptx")

			'Comment author
			Dim author As ICommentAuthor = presentation.CommentAuthors.AddAuthor("E-iceblue", "comment:")

			'Add comment
			presentation.Slides(0).AddComment(author, "Add comment", New PointF(18, 25), Date.Now)

			'Save the document
			presentation.SaveToFile("AddComment.pptx", FileFormat.Pptx2010)
			Process.Start("AddComment.pptx")
		End Sub
	End Class
End Namespace
