Imports Spire.Presentation
Imports Spire.Presentation.Charts
Imports Spire.Presentation.Drawing.Animation

Namespace ApplyAnimationOnChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\AnimationChart.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Get chart
			Dim shape As IShape = slide.Shapes(0)
			If TypeOf shape Is IChart Then
				'Apply Wipe animation effect to the chart
				Dim effect As AnimationEffect = slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Wipe)
				'Set the BuildType as Series
				effect.GraphicAnimation.BuildType = GraphicBuildType.BuildAsSeries
			End If

			'Save the document
			Dim result As String = "ApplyAnimationOnChart.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)

			'Launch the document
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