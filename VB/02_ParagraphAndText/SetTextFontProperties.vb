Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace SetTextFontProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Add a new shape to the PPT document
			Dim rec As New RectangleF(presentation.SlideSize.Size.Width \ 2 - 250, 80, 500, 150)
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, rec)

			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None

			'Add text to the shape
			shape.AppendTextFrame("Welcome to use Spire.Presentation")

			Dim textRange As TextRange = shape.TextFrame.TextRange
			'Set the font
			textRange.LatinFont = New TextFont("Times New Roman")
			'Set bold property of the font
			textRange.IsBold = TriState.True

			'Set italic property of the font
			textRange.IsItalic = TriState.True

			'Set underline property of the font
			textRange.TextUnderlineType = TextUnderlineType.Single

			'Set the height of the font
			textRange.FontHeight = 50

			'Set the color of the font
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.CadetBlue

			Dim result As String = "SetTextFontProperties_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)
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