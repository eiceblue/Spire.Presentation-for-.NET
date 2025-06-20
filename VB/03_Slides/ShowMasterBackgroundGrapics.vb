Imports Spire.Presentation

Namespace ShowMasterBackgroundGraphics
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a Presentation object and load the input file 
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ShowMasterBackgroundGraphics.pptx")

			' Set whether to show the background graphics of the slide master
			presentation.Slides(0).Layout.ShowMasterShapes = True

			' Save file 
			presentation.SaveToFile("ShowMasterBackgroundGrapics_output.pptx",FileFormat.Pptx2019)

			'Dispose
			presentation.Dispose()

			Process.Start("ShowMasterBackgroundGrapics_output.pptx")
		End Sub
	End Class
End Namespace