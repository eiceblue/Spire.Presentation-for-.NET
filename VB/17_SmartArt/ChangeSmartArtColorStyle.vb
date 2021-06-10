Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace ChangeSmartArtColorStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddSmartArtNode.pptx")

			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ISmartArt Then
					'Get the SmartArt and collect nodes
					Dim smartArt As ISmartArt = TryCast(shape, ISmartArt)
					' Check SmartArt color type
					If smartArt.ColorStyle = SmartArtColorType.ColoredFillAccent1 Then
						' Change SmartArt color type
						smartArt.ColorStyle = SmartArtColorType.ColorfulAccentColors
					End If
				End If
			Next shape
			Dim result As String = "ChangeSmartArtColorStyle_result.pptx"
			'Save the file
			presentation.SaveToFile(result, FileFormat.Pptx2010)

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