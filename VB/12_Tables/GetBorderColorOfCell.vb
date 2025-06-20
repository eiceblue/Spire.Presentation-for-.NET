Imports System.IO
Imports System.Text
Imports Spire.Presentation

Namespace GetBorderColorOfCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a PowerPoint document
			Dim presentation As New Presentation()

			'Load file from disk
			presentation.LoadFromFile("..\..\..\..\..\..\Data\GetBorderColorOfCell.pptx")

			'Get the table in the first slide
			Dim table As ITable = TryCast(presentation.Slides(0).Shapes(0), ITable)

			'Get borders' color of the first cell
			Dim sb As New StringBuilder()
            sb.Append("Color of left border:")
            sb.Append(table(0, 0).BorderLeftDisplayColor)
            sb.AppendLine()
            sb.Append("Color of top border:")
            sb.Append(table(0, 0).BorderTopDisplayColor)
            sb.AppendLine()
            sb.Append("Color of right border:")
            sb.Append(table(0, 0).BorderRightDisplayColor)
            sb.AppendLine()
            sb.Append("Color of bottom border:")
            sb.Append(table(0, 0).BorderBottomDisplayColor)
            sb.AppendLine()
			'Get display color of the first cell
            sb.Append("Color of cell:")
            sb.Append(table(0, 0).DisplayColor)
			Dim result As String = "Result-SetChartDataLabelRange.txt"

			File.WriteAllText(result, sb.ToString())
			Process.Start(result)
		End Sub


	End Class
End Namespace