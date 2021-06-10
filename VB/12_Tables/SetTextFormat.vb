Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation
Imports Spire.Presentation.Drawing.Transition
Imports Spire.Presentation.Diagrams
Imports System.IO

Namespace SetTextFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load PPT file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\Table.pptx")
			'Get the first slide
			Dim slide As ISlide = presentation.Slides(0)
			Dim str As New StringBuilder()
			For Each shape As IShape In slide.Shapes
				'Verify if it is table
				If TypeOf shape Is ITable Then
					Dim table As ITable = CType(shape, ITable)

					Dim cell1 As Cell = table.TableRows(0)(0)
					'Set table cell's text alignment type 
					cell1.TextAnchorType = TextAnchorType.Top
					'Set italic style
					cell1.TextFrame.TextRange.Format.IsItalic = TriState.True

					Dim cell2 As Cell = table.TableRows(1)(0)
					'Set table cell's foreground color
					cell2.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					cell2.TextFrame.TextRange.Fill.SolidColor.Color = Color.Green
					'Set table cell's background color
					cell2.FillFormat.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					cell2.FillFormat.SolidColor.Color = Color.LightGray


					Dim cell3 As Cell = table.TableRows(2)(2)
					'Set table cell's font and font size
					cell3.TextFrame.TextRange.FontHeight = 12
					cell3.TextFrame.TextRange.LatinFont = New TextFont("Arial Black")
					cell3.TextFrame.TextRange.HighlightColor.Color = Color.YellowGreen


					Dim cell4 As Cell = table.TableRows(2)(1)
					'Set table cell's margin and borders
					cell4.MarginLeft = 20
					cell4.MarginTop = 30
					cell4.BorderTop.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					cell4.BorderTop.SolidFillColor.Color = Color.Red
					cell4.BorderBottom.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					cell4.BorderBottom.SolidFillColor.Color = Color.Red
					cell4.BorderLeft.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					cell4.BorderLeft.SolidFillColor.Color = Color.Red
					cell4.BorderRight.FillType = Spire.Presentation.Drawing.FillFormatType.Solid
					cell4.BorderRight.SolidFillColor.Color = Color.Red
				End If
			Next shape

			Dim result As String = "SetTextFormat_result.pptx"
			presentation.SaveToFile(result, FileFormat.Pptx2013)
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