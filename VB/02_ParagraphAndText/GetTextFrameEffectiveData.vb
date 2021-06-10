Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO
Imports Spire.Presentation.Drawing

Namespace GetTextFrameEffectiveData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Template_Az1.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Get a shape 
			Dim shape As IAutoShape = TryCast(presentation.Slides(0).Shapes(0), IAutoShape)

			Dim textFrameFormat As ITextFrameProperties = shape.TextFrame
			Dim str As New StringBuilder()
			str.AppendLine("Anchoring type: " & textFrameFormat.AnchoringType)
			str.AppendLine("Autofit type: " & textFrameFormat.AutofitType)
			str.AppendLine("Text vertical type: " & textFrameFormat.VerticalTextType)
			str.AppendLine("Margins")
			str.AppendLine("   Left: " & textFrameFormat.MarginLeft)
			str.AppendLine("   Top: " & textFrameFormat.MarginTop)
			str.AppendLine("   Right: " & textFrameFormat.MarginRight)
			str.AppendLine("   Bottom: " & textFrameFormat.MarginBottom)

			Dim result As String = "GetTextFrameEffectiveData_result.txt"
			File.WriteAllText(result, str.ToString())
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