Imports Spire.Presentation

Namespace ReplaceAndFormatText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Presentation object.
			Dim ppt As New Presentation()

			' Load a PowerPoint presentation from the specified file.
			ppt.LoadFromFile("..\..\..\..\..\..\Data\TextTemplate.pptx")

			' Create a new object to store the default text range formatting properties.
			Dim format As New DefaultTextRangeProperties()

			' Set the IsBold property of the text range formatting to true, making the text bold.
			format.IsBold = TriState.True

			' Set the FillType property of the text range fill to Solid, indicating a solid fill color.
			format.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid

			' Set the Color property of the solid fill color to red.
			format.Fill.SolidColor.Color = Color.Red

			' Set the FontHeight property of the text range formatting to 25, indicating the font size.
			format.FontHeight = 25

			' Replace all occurrences of the text "Spire.Presentation for .NET" with "Spire.PPT" and apply the specified formatting.
			ppt.ReplaceAndFormatText("Spire.Presentation for .NET", "Spire.PPT", format)

			' Specify the name for the output PowerPoint presentation file.
			Dim result As String = "output.pptx"

			' Save the modified presentation to the specified output file in the PPTX format compatible with PowerPoint 2016.
			ppt.SaveToFile(result, FileFormat.Pptx2016)

			' Dispose of the Presentation object to free up resources
			ppt.Dispose()

			' Launch the saved file
			PresentationDocViewer(result)
		End Sub

		Private Sub PresentationDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace