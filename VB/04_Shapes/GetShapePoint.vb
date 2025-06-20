Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Spire.Presentation

Namespace GetShapePoint
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a PPT document
			Dim ppt As New Presentation()
			ppt.LoadFromFile("../../../../../../Data/ShapePoint.pptx")

			'Get the first shape in first slide
			Dim shape As IAutoShape = CType(ppt.Slides(0).Shapes(0), IAutoShape)

			'Get the Point of shape
			Dim points As IList(Of PointF) = shape.Points

			Dim sb As New StringBuilder()
			sb.Append("point count£º" & " " & points.Count & vbCrLf)

			For i As Integer = 0 To points.Count - 1
				sb.Append("point" & i & " " & points(i).ToString() & vbCrLf)
			Next i

			'Save the result txt file           
			File.WriteAllText("PointInformation.txt", sb.ToString())
			Process.Start("PointInformation.txt")
		End Sub
	End Class
End Namespace