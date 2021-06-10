Imports Spire.Presentation
Imports Spire.Presentation.Diagrams

Namespace AddHyperlinkToSmartArtNode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create an instance of presentation document
			Dim ppt As New Presentation()
			'Load file
			ppt.LoadFromFile("..\..\..\..\..\..\Data\SmartArtNode.pptx")

			'Get the smartArt shape
			Dim sr As ISmartArt = TryCast(ppt.Slides(0).Shapes(0), ISmartArt)
			'Add hylerlinks to the nodes
			Dim node As ISmartArtNode = sr.Nodes(0)
			node.Click = New ClickHyperlink(ppt.Slides(1))
			node = sr.Nodes(1)
			node.Click = New ClickHyperlink(ppt.Slides(2))
			node = sr.Nodes(2)
			node.Click = New ClickHyperlink(ppt.Slides(3))
			'Save the document
			Dim result As String = "AddHyperlinkToSmartArtNode.pptx"
			ppt.SaveToFile(result, FileFormat.Pptx2013)
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