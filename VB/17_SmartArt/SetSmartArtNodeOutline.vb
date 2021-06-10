Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Diagrams
Imports Spire.Presentation.Drawing

Namespace SetSmartArtNodeOutline
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim ppt As New Presentation()
			'Load the document from disk
			ppt.LoadFromFile("..\..\..\..\..\..\Data\CreateSmartArtShape.pptx")
			'Set ISmartArt form special shape
			Dim smartArt As ISmartArt = TryCast(ppt.Slides(0).Shapes(0), ISmartArt)
			Dim count As Integer = smartArt.Nodes.Count
			Dim node As ISmartArtNode
			'Loop through all nodes
			For i As Integer = 0 To count - 1
				node = smartArt.Nodes(i)
				'Set the fill format type
				node.Line.FillType = FillFormatType.Solid
				'Set the line style
				node.Line.Style = TextLineStyle.ThinThin
				'Set the line color
				node.Line.SolidFillColor.Color = Color.Red
				'Set the line width
				node.Line.Width = 2
			Next i
			'Save the document
			Dim result As String = "SetSmartArtNodeOutline.pptx"
			ppt.SaveToFile(result, Spire.Presentation.FileFormat.Pptx2013)
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