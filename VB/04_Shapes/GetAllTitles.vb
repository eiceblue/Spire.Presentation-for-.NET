Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace GetAllTitles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\Titles.pptx")

			'Instantiate a list of IShape objects
			Dim shapelist As New List(Of IShape)()
			'Loop through all sildes and all shapes on each slide
			For Each slide As ISlide In ppt.Slides
				For Each shape As IShape In slide.Shapes
					If shape.Placeholder IsNot Nothing Then
						'Get all titles
						Select Case shape.Placeholder.Type
							Case PlaceholderType.Title
								shapelist.Add(shape)
							Case PlaceholderType.CenteredTitle
								shapelist.Add(shape)
							Case PlaceholderType.Subtitle
								shapelist.Add(shape)
						End Select
					End If
				Next shape
			Next slide

			'Loop through the list and get the inner text of all shapes in the list
			Dim sb As New StringBuilder()
			sb.AppendLine("Below are all the obtained titles:")
			For i As Integer = 0 To shapelist.Count - 1
				Dim shape1 As IAutoShape = TryCast(shapelist(i), IAutoShape)
				sb.AppendLine(shape1.TextFrame.Text)
			Next i

			'Save to the Text file
			Dim result As String = "GetAllTitles.txt"
			File.WriteAllText(result, sb.ToString())
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