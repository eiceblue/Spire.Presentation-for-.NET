Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing
Imports System.IO
Imports Spire.Presentation.Charts

Namespace AddAndDetectMathEquations
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create Presentation
			Dim presentation As New Presentation()

			'Math code
			Dim latexMathCode As String = "x^{2}+\sqrt{x^{2}+1}=2"

			'Append a shape
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, New RectangleF(30, 100, 400, 30))
			shape.TextFrame.Paragraphs.Clear()

			'Add math equation
			Dim tp As TextParagraph = shape.TextFrame.Paragraphs.AddParagraphFromLatexMathCode(latexMathCode)

			'Detect if the slide contains math equation
			For i As Integer = 0 To presentation.Slides(0).Shapes.Count - 1

				If TypeOf presentation.Slides(0).Shapes(i) Is IAutoShape Then
					Dim containMathEquation As Boolean = (TryCast(presentation.Slides(0).Shapes(i), IAutoShape)).ContainMathEquation
					MessageBox.Show("The first slide contains math equations: " & containMathEquation)
				End If
			Next i

			'Save the file
			Dim result As String = "AddAndDetectMathEquations_result.pptx"
			presentation.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013)

			'Launching the result file.
			Viewer(result)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace