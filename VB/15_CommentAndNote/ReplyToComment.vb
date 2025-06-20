Imports Spire.Presentation

Namespace ReplyToComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create ppt file
			Dim ppt As New Presentation()

			'Create Comment author
			Dim author As ICommentAuthor = ppt.CommentAuthors.AddAuthor("E-iceblue", "comment")

			'Add comment
			ppt.Slides(0).AddComment(author, "Add comment", New Point(18, 25), Date.Now)
			Dim comment As Comment = ppt.Slides(0).Comments(0)

			'Add reply to Comment
			If Not comment.IsReply Then
				comment.Reply(author, "Add Reply1", Date.Now)
				comment.Reply(author, "Add Reply2", Date.Now)
			End If

			'delete first reply
			ppt.Slides(0).DeleteComment(author, "Add Reply1")

			'Save the result ppt file
			ppt.SaveToFile("AddReplyToComment.pptx", FileFormat.Pptx2013)
			Process.Start("AddReplyToComment.pptx")
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace