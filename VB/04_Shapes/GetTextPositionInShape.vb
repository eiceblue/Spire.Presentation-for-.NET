Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace GetTextPositionInShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Presentation object
			Dim ppt As New Presentation()

			' Load a PowerPoint file from a specified location
			ppt.LoadFromFile("..\..\..\..\..\..\Data\GetTextPositionInShape.pptx")

			' Create a StringBuilder to store text information
			Dim sb As New StringBuilder()

			' Access the first slide in the presentation
			Dim slide As ISlide = ppt.Slides(0)

			' Iterate through all the shapes in the slide
			For i As Integer = 0 To slide.Shapes.Count - 1
				' Get the current shape
				Dim shape As IShape = slide.Shapes(i)

				' Check if the shape is an AutoShape
				If TypeOf shape Is IAutoShape Then
					' Cast the shape to an AutoShape
					Dim autoshape As IAutoShape = TryCast(slide.Shapes(i), IAutoShape)

					' Get the text content of the AutoShape
					Dim text As String = autoshape.TextFrame.Text

					' Obtain the text position information within the AutoShape
					Dim point As PointF = autoshape.TextFrame.GetTextLocation()

					' Append information about the shape, text, and location to the StringBuilder
					sb.AppendLine("Shape " & i & "£º" & text & vbCrLf & "location£º" & point.ToString())
				End If
			Next i
			' Specify the name for the result file
			Dim result As String = "GetTextPositionInShape.txt"

			' Append the collected information to a text file named "GetTextPositionInShape.txt"
			File.AppendAllText(result, sb.ToString())

			' Dispose of the Presentation object to release resources
			ppt.Dispose()

			' Launch the result file 
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