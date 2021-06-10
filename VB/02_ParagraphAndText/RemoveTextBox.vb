Imports Spire.Presentation

Namespace RemoveTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\TextBoxTemplate.pptx")

			'Get the first slide
			Dim slide As ISlide = ppt.Slides(0)
			'Traverse all the shapes in slide
			Dim i As Integer = 0
			Do While (i < slide.Shapes.Count)
				If slide.Shapes(i).Name.Contains("TextBox") Then
					slide.Shapes.RemoveAt(i)
					i=(i-1)
				End If
				i=(i+1)
			Loop

			'Save the document
			Dim result As String = "RemoveTextBox.pptx"
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