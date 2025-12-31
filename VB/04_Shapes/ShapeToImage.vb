Imports Spire.Presentation
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace ShapeToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			'Create a PPT document
			Dim presentation As New Presentation()
			presentation.LoadFromFile("..\..\..\..\..\..\Data\ShapeToImage.pptx")

			For i As Integer = 0 To presentation.Slides(0).Shapes.Count - 1
				Dim fileName As String = String.Format("Picture-{0}.png", i)
				'Save shapes as images
				Dim image As Image = presentation.Slides(0).Shapes(i).SaveAsImage()

				'The following method also can save shape as image
				'Image image = presentation.Slides[0].Shapes.SaveAsImage(i);

				'Write image to Png
				image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png)
				System.Diagnostics.Process.Start(fileName)
			Next i

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'             for (int i = 0; i < presentation.Slides[0].Shapes.Count; i++)
'            {
'                using (var images = presentation.Slides[0].Shapes.SaveAsImage(i))
'                {
'                    string filename = String.Format("ShapeToImage-{0}.png", i);
'                    FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
'                    images.CopyTo(fileStream);
'                    fileStream.Flush();
'                    images.Dispose();
'                }
'            }
'            

			presentation.Dispose()
		End Sub
	End Class
End Namespace
