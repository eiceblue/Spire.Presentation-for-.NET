Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Spire.Presentation.Drawing

Public Class Form1

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        'create PPT document
        Dim presentation As New Presentation()

        'set background Image
        Dim ImageFile As String = "..\..\..\..\..\..\Data\bg.png"
        Dim rect As New RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
        presentation.Slides(0).Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect)
        presentation.Slides(0).Shapes(0).Line.FillFormat.SolidFillColor.Color = Color.FloralWhite

        Dim widths As [Double]() = New Double() {100, 100, 150, 100, 100}
        Dim heights As [Double]() = New Double() {15, 15, 15, 15, 15, 15, _
         15, 15, 15, 15, 15, 15, _
         15}

        'add new table to PPT
        Dim table As ITable = presentation.Slides(0).Shapes.AppendTable(presentation.SlideSize.Size.Width / 2 - 275, 90, widths, heights)

        Dim dataStr As [String](,) = New [String](,) {{"Name", "Capital", "Continent", "Area", "Population"}, {"Venezuela", "Caracas", "South America", "912047", "19700000"}, {"Bolivia", "La Paz", "South America", "1098575", "7300000"}, {"Brazil", "Brasilia", "South America", "8511196", "150400000"}, {"Canada", "Ottawa", "North America", "9976147", "26500000"}, {"Chile", "Santiago", "South America", "756943", "13200000"}, _
         {"Colombia", "Bagota", "South America", "1138907", "33000000"}, {"Cuba", "Havana", "North America", "114524", "10600000"}, {"Ecuador", "Quito", "South America", "455502", "10600000"}, {"Paraguay", "Asuncion", "South America", "406576", "4660000"}, {"Peru", "Lima", "South America", "1285215", "21600000"}, {"Jamaica", "Kingston", "North America", "11424", "2500000"}, _
         {"Mexico", "Mexico City", "North America", "1967180", "88600000"}}

        'add data to table
        For i As Integer = 0 To 12
            For j As Integer = 0 To 4
                'fill the table with data
                table(j, i).TextFrame.Text = dataStr(i, j)

                'set the Font
                table(j, i).TextFrame.Paragraphs(0).TextRanges(0).LatinFont = New TextFont("Arial Narrow")
            Next
        Next

        'set the alignment of the first row to Center
        For i As Integer = 0 To 4
            table(i, 0).TextFrame.Paragraphs(0).Alignment = TextAlignmentType.Center
        Next

        'set the style of table
        table.StylePreset = TableStylePreset.LightStyle3Accent1

        'save the document
        presentation.SaveToFile("table.pptx", FileFormat.Pptx2010)
        System.Diagnostics.Process.Start("table.pptx")

    End Sub
End Class