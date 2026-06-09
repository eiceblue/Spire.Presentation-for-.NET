Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Spire.Presentation

Namespace ToImageWithResolution
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create an instance of presentation document
            Dim ppt As Presentation = New Presentation()
           
            'Load file
            ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Conversion.pptx")

            ' Render the first slide to an image with custom DPI
            Dim img As Image = ppt.SaveToImage(0,100,100)

            'Save image to file
            Dim result As String = "ToImageWithResolution.png"
            img.Save(result, System.Drawing.Imaging.ImageFormat.Png)

            'Launch the PPT file
            System.Diagnostics.Process.Start(result)

        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace