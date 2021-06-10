Imports System.ComponentModel
Imports System.Text
Imports Spire.Presentation.Drawing
Imports Spire.Presentation

Namespace CreateTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PPT document
			Dim presentation As New Presentation()

			'Load the document from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\CreateTable.pptx")

			Dim widths() As Double = { 100, 100, 150, 100, 100 }
			Dim heights() As Double = { 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15 }

			'Add new table to PPT
			Dim table As ITable = presentation.Slides(0).Shapes.AppendTable(presentation.SlideSize.Size.Width \ 2 - 275, 90, widths, heights)

			Dim dataStr(,) As String = { {"Name", "Capital", "Continent", "Area", "Population"}, {"Venezuela", "Caracas", "South America", "912047", "19700000"}, {"Bolivia", "La Paz", "South America", "1098575", "7300000"}, {"Brazil", "Brasilia", "South America", "8511196", "150400000"}, {"Canada", "Ottawa", "North America", "9976147", "26500000"}, {"Chile", "Santiago", "South America", "756943", "13200000"}, {"Colombia", "Bagota", "South America", "1138907", "33000000"}, {"Cuba", "Havana", "North America", "114524", "10600000"}, {"Ecuador", "Quito", "South America", "455502", "10600000"}, {"Paraguay", "Asuncion","South America", "406576", "4660000"}, {"Peru", "Lima", "South America", "1285215", "21600000"}, {"Jamaica", "Kingston", "North America", "11424", "2500000"}, {"Mexico", "Mexico City", "North America", "1967180", "88600000"} }

			'Add data to table
			For i As Integer = 0 To 12
				For j As Integer = 0 To 4
					'Fill the table with data
					table(j, i).TextFrame.Text = dataStr(i, j)

					'Set the Font
					table(j, i).TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Narrow")
				Next j
			Next i

			'Set the alignment of the first row to Center
			For i As Integer = 0 To 4
				table(i, 0).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Center
			Next i

			'Set the style of table
			table.StylePreset = TableStylePreset.LightStyle3Accent1

			'Save the document
			presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010)
			Process.Start("Output.pptx")

		End Sub
	End Class
End Namespace