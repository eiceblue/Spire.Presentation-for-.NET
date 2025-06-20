Imports Spire.Presentation

Namespace InsertPlaceholderInMaster
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a Presentation object
			Dim presentation As New Presentation()

			 ' Inset palce hodler 
			presentation.Masters(0).Layouts(0).InsertPlaceholder(InsertPlaceholderType.Text, New RectangleF(20, 30, 400, 400))

			' Save file 
			presentation.SaveToFile("InsertPlaceholderInMaster_output.pptx", FileFormat.Pptx2019)

			'Dispose
			presentation.Dispose()

			Process.Start("InsertPlaceholderInMaster_output.pptx")
		End Sub
	End Class
End Namespace