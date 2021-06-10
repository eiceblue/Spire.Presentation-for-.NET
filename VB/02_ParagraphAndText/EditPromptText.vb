Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace EditPromptText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim loadPath As String = "..\..\..\..\..\..\Data\HasPromptText.pptx"
			Dim savePath As String = "EditPromptText.pptx"
			'Load a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile(loadPath)

			' Iterate through the slide
			For Each shape As IShape In presentation.Slides(0).Shapes
				If shape.Placeholder IsNot Nothing AndAlso TypeOf shape Is IAutoShape Then
					Dim text As String = ""
					' Set the text of the title
					If shape.Placeholder.Type = PlaceholderType.CenteredTitle Then
						text = "custom title create by Spire"
					' Set text of the subtitle.
					ElseIf shape.Placeholder.Type = PlaceholderType.Subtitle Then
						text = "custom subtitle create by Spire"
					End If

					TryCast(shape, IAutoShape).TextFrame.Text = text
				End If
			Next shape

			'Save the file
			presentation.SaveToFile(savePath, FileFormat.Pptx2013)
			Process.Start(savePath)
		End Sub
	End Class
End Namespace