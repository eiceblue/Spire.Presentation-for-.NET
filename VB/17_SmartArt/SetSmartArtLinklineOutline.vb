Imports System.Collections
Imports Spire.Presentation
Imports Spire.Presentation.Diagrams
Imports Spire.Presentation.Drawing

Namespace SetSmartArtLinklineOutline
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
			'Get the specified shape as ISmartArt
			Dim smartArt As ISmartArt = TryCast(ppt.Slides(0).Shapes(0), ISmartArt)
			Dim count As Integer = smartArt.Nodes.Count
			Dim node As ISmartArtNode
			'Loop through all smartArts
			For i As Integer = 0 To count - 1
				node = smartArt.Nodes(i)
				'Set the line type
				node.LinkLine.FillType = FillFormatType.Solid
				'Set the line color
				node.LinkLine.SolidFillColor.Color = Color.Red
				'Set the line width
				node.LinkLine.Width = 2
				'Set the line DashStyle
				node.LinkLine.DashStyle = LineDashStyleType.SystemDash
			Next i
			'Save the document
			Dim result As String = "SetSmartArtLinklineOutline.pptx"
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