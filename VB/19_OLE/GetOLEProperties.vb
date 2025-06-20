Imports System.Collections
Imports System.IO
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Collections

Namespace GetOLEProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GetOLEPropertiesOutsideOfShape.pptx")

			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)

			'Get the first OLE
			Dim oles As OleObjectCollection = slide.OleObjects
			Dim oleObject As OleObject = oles(0)

			Dim sb As New StringBuilder()

			'Get the information of OLE Object
			sb.AppendLine("ShapeID=" & oleObject.ShapeID)
			sb.AppendLine("FrameTop=" & oleObject.Frame.Top)
			sb.AppendLine("FrameLeft=" & oleObject.Frame.Left)
			sb.AppendLine("FrameWidth=" & oleObject.Frame.Width)
			sb.AppendLine("FrameHight=" & oleObject.Frame.Height)
			sb.AppendLine("IsHidden=" & oleObject.IsHidden)

			'Get the properties of OLE
			For Each entry As DictionaryEntry In oleObject.Properties
				sb.AppendLine(entry.Key & ":" & entry.Value)
			Next entry

			' Save and preview the output file
			File.AppendAllText("GetOLEOutsideOfShape.txt", sb.ToString())

			Process.Start("GetOLEOutsideOfShape.txt")

			presentation.Dispose()
		End Sub
	End Class
End Namespace