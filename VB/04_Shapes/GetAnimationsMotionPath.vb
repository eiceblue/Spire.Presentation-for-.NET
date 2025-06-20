Imports System.IO
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Animation

Namespace GetAnimationsMotionPath
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document
			Dim presentation As New Presentation()
			'Load the file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\..\Data\GetAnimationsMotionPath.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			'Get the first shape
			Dim shape As IShape = slide.Shapes(0)
			'Create a StringBuilder to save the tracks
			Dim sb As New StringBuilder()
			Dim i As Integer = 1
			'Traverse all animations
			For Each effect As AnimationEffect In shape.Slide.Timeline.MainSequence
				If effect.ShapeTarget.Equals(TryCast(shape, Shape)) Then
					'Get MotionPath
					Dim path As MotionPath = (CType(effect.CommonBehaviorCollection(0), AnimationMotion)).Path
					'Get all points in the path
					For Each motionCmdPath As MotionCmdPath In path
						Dim points() As PointF = motionCmdPath.Points
						Dim type As MotionCommandPathType = motionCmdPath.CommandType
						If points IsNot Nothing Then
							For Each point As PointF In points
								sb.AppendLine(i & "  MotionType: " & type & " -> X: " & point.X & ", Y: " & point.Y)
							Next point
							i += 1
						End If
					Next motionCmdPath
				End If
			Next effect
			Dim result As String = "GetAnimationsMotionPath.txt"
			File.WriteAllText(result, sb.ToString())
			Process.Start(result)
		End Sub
	End Class
End Namespace