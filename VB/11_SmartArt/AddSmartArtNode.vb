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
			'create PPT document
			Dim presentation As New Presentation()

			'load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SmartArt.pptx")

			'get the SmartArt
			Dim sa As ISmartArt = TryCast(presentation.Slides(0).Shapes(0), ISmartArt)

			'add a node
			Dim node As ISmartArtNode = sa.Nodes.AddNode()
			'add text and set the text style 
			node.TextFrame.Text = "AddText"
			node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink


			presentation.SaveToFile("AddNode.pptx", FileFormat.Pptx2010)
			Process.Start("AddNode.pptx")
		End Sub
	End Class
End Namespace
