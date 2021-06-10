Imports Spire.Presentation
Imports Spire.Presentation.Diagrams
Imports System.ComponentModel
Imports System.Text

Namespace RemoveNode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\RemoveNode.pptx")

			'Get the SmartArt and collect nodes
			Dim sa As ISmartArt = TryCast(presentation.Slides(0).Shapes(0), ISmartArt)
			Dim nodes As ISmartArtNodeCollection = sa.Nodes

			'Remove the node to specific position
			nodes.RemoveNodeByPosition(2)

			presentation.SaveToFile("RemoveNode.pptx", FileFormat.Pptx2010)
			Process.Start("RemoveNode.pptx")
		End Sub
	End Class
End Namespace
