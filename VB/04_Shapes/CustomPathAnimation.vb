Imports Spire.Presentation
Imports Spire.Presentation.Collections
Imports Spire.Presentation.Drawing.Animation

Namespace CustomPathAnimation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim ppt As New Presentation()

			'Add shape
			Dim shape As IAutoShape = ppt.Slides(0).Shapes.AppendShape(ShapeType.Rectangle, New RectangleF(0, 0, 200, 200))

			'Add animation
			Dim effect As AnimationEffect = ppt.Slides(0).Timeline.MainSequence.AddEffect(shape, AnimationEffectType.PathUser)
			Dim common As CommonBehaviorCollection = effect.CommonBehaviorCollection
			Dim motion As AnimationMotion = CType(common(0), AnimationMotion)
			motion.Origin = AnimationMotionOrigin.Layout
			motion.PathEditMode = AnimationMotionPathEditMode.Relative

			'Add moin path
			Dim moinPath As New MotionPath()
			moinPath.Add(MotionCommandPathType.MoveTo, New PointF() { New PointF(0, 0) }, MotionPathPointsType.CurveAuto, True)
			moinPath.Add(MotionCommandPathType.LineTo, New PointF() { New PointF(0.1f, 0.1f) }, MotionPathPointsType.CurveAuto, True)
			moinPath.Add(MotionCommandPathType.LineTo, New PointF() { New PointF(-0.1f, 0.2f) }, MotionPathPointsType.CurveAuto, True)
			moinPath.Add(MotionCommandPathType.End, New PointF() { }, MotionPathPointsType.CurveStraight, True)
			motion.Path = moinPath

			'Save the document
			Dim outputFile As String = "result.pptx"
			ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
			ppt.Dispose()

			'Launch the PPT file
			FileViewer(outputFile)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
