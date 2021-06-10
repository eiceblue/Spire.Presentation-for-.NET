Imports Spire.Presentation
Imports System.ComponentModel
Imports System.Text

Namespace CreateSmartArtShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\CreateSmartArtShape.pptx")

			Dim sa As Spire.Presentation.Diagrams.ISmartArt = presentation.Slides(0).Shapes.AppendSmartArt(200, 60, 300, 300, Spire.Presentation.Diagrams.SmartArtLayoutType.Gear)

			'Set type and color of smartart
			sa.Style = Spire.Presentation.Diagrams.SmartArtStyleType.SubtleEffect
			sa.ColorStyle = Spire.Presentation.Diagrams.SmartArtColorType.GradientLoopAccent3

			'Remove all shapes
			For Each a As Object In sa.Nodes
				sa.Nodes.RemoveNode(0)
			Next a

			'Add two custom shapes with text
			Dim node As Spire.Presentation.Diagrams.ISmartArtNode = sa.Nodes.AddNode()
			sa.Nodes(0).TextFrame.Text = "aa"
			node = sa.Nodes.AddNode()
			node.TextFrame.Text = "bb"
			node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
			node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Black

			'Save and launch the file
			presentation.SaveToFile("CreateSmartArtShape.pptx", FileFormat.Pptx2010)
			Process.Start("CreateSmartArtShape.pptx")
		End Sub
	End Class
End Namespace
