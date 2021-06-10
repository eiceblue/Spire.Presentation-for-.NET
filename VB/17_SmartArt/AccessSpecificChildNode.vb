Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace AccessSpecificChildNode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create PPT document
			Dim presentation As New Presentation()

			'Load the PPT
			presentation.LoadFromFile("..\..\..\..\..\..\Data\SmartArt.pptx")

			Dim strB As New StringBuilder()
			strB.AppendLine("Access SmartArt child node at specific position.")
			strB.AppendLine("Here is the SmartArt child node parameters details:")
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ISmartArt Then
					'Get the SmartArt
					Dim sa As ISmartArt = TryCast(shape, ISmartArt)

					'Get SmartArt node collection 
					Dim nodes As ISmartArtNodeCollection = sa.Nodes

					'Access SmartArt node at index 0
					Dim node As ISmartArtNode = nodes(0)

					'Access SmartArt child node at index 1
					Dim childNode As ISmartArtNode = node.ChildNodes(1)

					'Print the SmartArt child node parameters
					Dim outString As String = String.Format("Node text = {0}, Node level = {1}, Node Position = {2}", childNode.TextFrame.Text, childNode.Level, childNode.Position)

					strB.AppendLine(outString)
				End If

			Next shape
			Dim result As String = "AccessSpecificChildNode_result.txt"
			'Save the file
			File.WriteAllText(result, strB.ToString())

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