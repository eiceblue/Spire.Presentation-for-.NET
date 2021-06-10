Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace AddNodeByPosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT
			presentation.LoadFromFile("..\..\..\..\..\..\Data\AddSmartArtNode2.pptx")

			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ISmartArt Then
					'Get the SmartArt and collect nodes
					Dim smartArt As ISmartArt = TryCast(shape, ISmartArt)

					Dim position As Integer = 0
					'Add a new node at specific position
					Dim node As ISmartArtNode = smartArt.Nodes.AddNodeByPosition(position)
					'Add text and set the text style 
					node.TextFrame.Text = "New Node"
					node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Red

					'Get a node
					node = smartArt.Nodes(1)
					position = 1
					'Add a new child node at specific position
					Dim childNode As ISmartArtNode = node.ChildNodes.AddNodeByPosition(position)
					'Add text and set the text style 
					node.TextFrame.Text = "New child node"
					node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Blue
				End If
			Next shape
			Dim result As String = "AddNodeByPosition_result.pptx"
			'Save the file
			presentation.SaveToFile(result, FileFormat.Pptx2010)

			Viewer(result)
		End Sub

		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace