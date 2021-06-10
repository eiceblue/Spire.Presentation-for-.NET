Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace LockAspectRatio
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Table.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			Dim str As New StringBuilder()
			For Each shape As IShape In slide.Shapes
				'Verify if it is table
				If TypeOf shape Is ITable Then
					Dim table As ITable = CType(shape, ITable)
					'Lock aspect ratio
					table.ShapeLocking.AspectRatioProtection = True
				End If
			Next shape

			Dim result As String = "LockAspectRatio_result.pptx"
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