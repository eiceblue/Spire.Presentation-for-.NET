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
			'create PPT document
			Dim presentation As New Presentation()

			'load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SmartArt.pptx")

			'get the SmartArt and collect nodes
			Dim sa As ISmartArt = TryCast(presentation.Slides(0).Shapes(0), ISmartArt)
			Dim nodes As ISmartArtNodeCollection = sa.Nodes

			'remove the node to specific position
			nodes.RemoveNodeByPosition(2)

			presentation.SaveToFile("RemoveNodes.pptx", FileFormat.Pptx2010)
			Process.Start("RemoveNodes.pptx")
		End Sub

		Private Sub lblDescription_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lblDescription.Click

		End Sub

		Private Sub pbLogo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pbLogo.Click

		End Sub
	End Class
End Namespace
