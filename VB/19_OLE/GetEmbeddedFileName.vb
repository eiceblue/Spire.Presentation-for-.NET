Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Data.SqlTypes
Imports System.IO
Imports System.Text


Namespace GetEmbeddedFileName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\oleTest.pptx")

			'Loop through the slides and shapes
			For Each slide As ISlide In presentation.Slides
				For Each shape As IShape In slide.Shapes
					If TypeOf shape Is IOleObject Then
						'Find OLE object
						Dim oleObject As IOleObject = TryCast(shape, IOleObject)
						'Get OLE object label name
						Dim oleFileName As String = oleObject.EmbeddedFileName
						MessageBox.Show("The name of the OLE object label is:" & oleFileName)
					End If
				Next shape
			Next slide



		End Sub
	End Class
End Namespace
