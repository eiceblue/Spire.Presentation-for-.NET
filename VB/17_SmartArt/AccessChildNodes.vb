Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace AccessChildNodes
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
			strB.AppendLine("Access SmartArt child nodes.")
			strB.AppendLine("Here is the SmartArt child node parameters details:")
			Dim outString As String = ""
			For Each shape As IShape In presentation.Slides(0).Shapes
				If TypeOf shape Is ISmartArt Then
					'Get the SmartArt and collect nodes
					Dim sa As ISmartArt = TryCast(shape, ISmartArt)
					Dim nodes As ISmartArtNodeCollection = sa.Nodes

					Dim position As Integer = 0
					'Access the parent node at position 0
					Dim node As ISmartArtNode = nodes(position)
					Dim childnode As ISmartArtNode
					'Traverse through all child nodes inside SmartArt
					For i As Integer = 0 To node.ChildNodes.Count - 1
						'Access SmartArt child node at index i
						childnode = node.ChildNodes(i)
						'Print the SmartArt child node parameters                       
						outString = String.Format("Node text = {0}, Node level = {1}, Node Position = {2}", childnode.TextFrame.Text, childnode.Level, childnode.Position)
						strB.AppendLine(outString)
					Next i

				End If

			Next shape
			Dim result As String = "AccessChildNode_result.txt"
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