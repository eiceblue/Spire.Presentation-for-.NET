Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Diagrams

Namespace AddSmartArtNode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddSmartArtNode.pptx")

			'Get the SmartArt
			Dim sa As ISmartArt = TryCast(presentation.Slides(0).Shapes(0), ISmartArt)

			'Add a node
			Dim node As ISmartArtNode = sa.Nodes.AddNode()
			'Add text and set the text style 
			node.TextFrame.Text = "AddText"
			node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink

			presentation.SaveToFile("AddSmartArtNode.pptx", FileFormat.Pptx2010)
			Process.Start("AddSmartArtNode.pptx")
		End Sub
	End Class
End Namespace
