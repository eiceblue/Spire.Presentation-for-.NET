Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing.Animation

Namespace SetAnimationRepeatType
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document and load file
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\Animation.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			Dim animations As AnimationEffectCollection = slide.Timeline.MainSequence

			animations(0).Timing.AnimationRepeatType = AnimationRepeatType.UtilEndOfSlide

			Dim result As String = "SetAnimationRepeatType_result.pptx"

			'Save to file.
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the PowerPoint file.
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