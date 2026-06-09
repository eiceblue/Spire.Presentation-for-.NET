Imports System
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace HighlightTextByRegex
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create a PPT document
            Dim presentation As Presentation = New Presentation()
            presentation.LoadFromFile(@"..\..\..\..\..\..\Data\SomePresentation.pptx")

            ' Create a regex pattern to match the word 
            Dim regex As Regex = New Regex(@"\bSpire\b")

            'Get the specified shape
            Dim shape As IAutoShape = (IAutoShape)presentation.Slides[0].Shapes[0]

            ' Initialize text highlighting options
            Dim options As TextHighLightingOptions = New TextHighLightingOptions()
            shape.TextFrame.HighLightRegex(regex, Color.Red, options)

            'Save to file
            Dim result As String = "HighlightTextByRegex_result.pptx"
            presentation.SaveToFile(result, FileFormat.Pptx2013)

            'Launch the PPT file
            System.Diagnostics.Process.Start(result)

        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace