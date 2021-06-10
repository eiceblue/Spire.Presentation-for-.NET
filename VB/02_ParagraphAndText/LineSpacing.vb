Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace LineSpacing
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Az.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Add a shape 
			Dim shape As IAutoShape = presentation.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(50, 100, presentation.SlideSize.Size.Width-100,300))
			shape.ShapeStyle.LineColor.Color = Color.White
			shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None
			shape.TextFrame.Paragraphs.Clear()

			'Add text
			shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers to" & "create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform." & "From Spire.Presentation v 3.7.5, Spire.Presentation starts to support .NET Core, .NET standard.")
			'Set font and color of text
			Dim textRange As TextRange = shape.TextFrame.TextRange
			textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			textRange.Fill.SolidColor.Color = Color.BlueViolet
			textRange.FontHeight =20
			textRange.LatinFont = New TextFont("Lucida Sans Unicode")

			'Set properties of paragraph
			shape.TextFrame.Paragraphs(0).SpaceBefore = 100
			shape.TextFrame.Paragraphs(0).SpaceAfter = 100
			shape.TextFrame.Paragraphs(0).LineSpacing = 150

			Dim result As String = "LineSpacing_result.pptx"
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