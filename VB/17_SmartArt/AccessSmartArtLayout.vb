Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace AccessSmartArtLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SmartArt.pptx")

			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ISmartArt Then
					'Get the SmartArt and collect nodes
					Dim sa As ISmartArt = TryCast(shape, ISmartArt)
					'Check SmartArt Layout
					Dim layout As String = sa.LayoutType.ToString()
					MessageBox.Show("SmartArt layout type is " & layout)
				End If
			Next shape
		End Sub
	End Class
End Namespace