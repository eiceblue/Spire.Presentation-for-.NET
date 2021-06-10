Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports System.IO

Namespace ExtractComments
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document.
			Dim presentation As New Presentation()

			'Load the file from disk.
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Ppt_5.pptx")

			Dim str As New StringBuilder()

			'Get all comments from the first slide.
			Dim comments() As Comment = presentation.Slides(0).Comments

			'Save the comments in txt file.
			For i As Integer = 0 To comments.Length - 1
				str.Append(comments(i).Text & vbCrLf)
			Next i

			Dim result As String = "Result-ExtractComments.txt"

			'Save to file.
			File.WriteAllText(result, str.ToString())

			'Launch the txt file.
			PptDocumentViewer(result)
		End Sub

		Private Sub PptDocumentViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace