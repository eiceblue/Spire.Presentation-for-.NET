Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation

Namespace AddMathMLEquation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()

			'Set the mathML code
			Dim mathMLCode As String = "<mml:math xmlns:mml=""http://www.w3.org/1998/Math/MathML"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">" & "<mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:msqrt><mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:msqrt><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:math>"

			'Add a shape
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New Rectangle(30, 100, 400, 30))
			shape.TextFrame.Paragraphs.Clear()

			'Add the mathml equation paragraph
			Dim tp As TextParagraph = shape.TextFrame.Paragraphs.AddParagraphFromMathMLCode(mathMLCode)

			'Save the document
			Dim outputFile As String = "result.pptx"
			ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
			ppt.Dispose()
			PresentationDocViewer(outputFile)
		End Sub
		Private Shared Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace