Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace GetSlideComments
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Comments.pptx")

			'Loop through comments
			For Each commentAuthor As ICommentAuthor In presentation.CommentAuthors
				For Each comment As Comment In commentAuthor.CommentsList
					'Get comment information
					Dim commentText As String = comment.Text
					Dim authorName As String = comment.AuthorName
					Dim time As Date = comment.DateTime
					MessageBox.Show("Comment text : " & comment.Text & vbLf & "Comment author : " & comment.AuthorName & vbLf & "Posted on time : " & comment.DateTime)
				Next comment
			Next commentAuthor
		End Sub
	End Class
End Namespace