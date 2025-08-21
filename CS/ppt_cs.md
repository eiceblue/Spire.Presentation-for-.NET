# Spire.Presentation C# Hello World
## Create a simple PowerPoint presentation with Hello World text
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add a new shape to the PPT document
RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 80, 500, 150);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);

shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

//Add text to the shape
shape.AppendTextFrame("Hello World!");

//Set the font and fill style of the text
TextRange textRange = shape.TextFrame.TextRange;
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
textRange.FontHeight = 66;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");
```

---

# Spire.Presentation CSharp Paragraph
## Add and format paragraph in presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a new shape
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 70, 620, 150));
shape.Fill.FillType = FillFormatType.None;
shape.ShapeStyle.LineColor.Color = Color.White;

//Set the alignment of paragraph
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
//Set the indent of paragraph
shape.TextFrame.Paragraphs[0].Indent = 50;
//Set the linespacing of paragraph
shape.TextFrame.Paragraphs[0].LineSpacing = 150;
//Set the text of paragraph
shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue.";

//Set the Font
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Rounded MT Bold");
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;
```

---

# spire.presentation csharp text alignment
## set paragraph alignment in powerpoint presentation
```csharp
//Get the related shape and set the text alignment
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[1];
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
shape.TextFrame.Paragraphs[1].Alignment = TextAlignmentType.Center;
shape.TextFrame.Paragraphs[2].Alignment = TextAlignmentType.Right;
shape.TextFrame.Paragraphs[3].Alignment = TextAlignmentType.Justify;
shape.TextFrame.Paragraphs[4].Alignment = TextAlignmentType.None;
```

---

# Spire.Presentation C# HTML
## Append HTML content to PowerPoint slides
```csharp
//Add a shape 
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 100, 200, 200));

//Clear default paragraphs 
shape.TextFrame.Paragraphs.Clear();

string code = "<html><body><p>This is a paragraph</p></body></html>";

//Append HTML, and generate a paragraph with default style in PPT document.
shape.TextFrame.Paragraphs.AddFromHtml(code);
string codeColor = "<html><body><p style=\" color:black \">This is a paragraph</p></body></html>";
//Append HTML with black setting
shape.TextFrame.Paragraphs.AddFromHtml(codeColor);

//Add another shape
IAutoShape shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(350, 100, 200, 200));

//Clear default paragraph 
shape1.TextFrame.Paragraphs.Clear();

//Change the fill format of shape
shape1.Fill.FillType = FillFormatType.Solid;
shape1.Fill.SolidColor.Color = Color.White;

//Append HTML
shape1.TextFrame.Paragraphs.AddFromHtml(code);
TextParagraph par = shape1.TextFrame.Paragraphs[0];
//Change the fill color for paragraph
foreach (TextRange tr in par.TextRanges)
{
    tr.Fill.FillType = FillFormatType.Solid;
    tr.Fill.SolidColor.Color = Color.Black;
}
```

---

# Spire.Presentation C# AutoFit Text or Shape
## Demonstrates how to set text autofit properties for shapes in a PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Set the AutofitType property to Shape
IAutoShape textShape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 100, 150, 80));
textShape2.TextFrame.Text = "Resize shape to fit text.";
textShape2.TextFrame.AutofitType = TextAutofitType.Shape;

//Set the AutofitType property to Normal
IAutoShape textShape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(400, 100, 150, 80));
textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape.";
textShape1.TextFrame.AutofitType = TextAutofitType.Normal;
```

---

# Spire.Presentation C# Borders and Shading
## Apply borders, gradient fill, and shadow effects to shapes in PowerPoint presentations
```csharp
// Get the shape from the presentation
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];

// Set line color and width of the border
shape.Line.FillType = FillFormatType.Solid;
shape.Line.Width = 3;
shape.Line.SolidFillColor.Color = Color.LightYellow;

// Set the gradient fill color of shape
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Gradient;
shape.Fill.Gradient.GradientShape = Spire.Presentation.Drawing.GradientShapeType.Linear;
shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.LightBlue);
shape.Fill.Gradient.GradientStops.Append(0, KnownColors.LightSkyBlue);

// Set the shadow for the shape
Spire.Presentation.Drawing.OuterShadowEffect shadow = new Spire.Presentation.Drawing.OuterShadowEffect();
shadow.BlurRadius = 20;
shadow.Direction = 30;
shadow.Distance = 8;
shadow.ColorFormat.Color = Color.LightSeaGreen;
shape.EffectDag.OuterShadowEffect = shadow;
```

---

# spire.presentation csharp bullets
## Add numbered bullets to paragraphs in PowerPoint presentation
```csharp
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[1];

foreach (TextParagraph para in shape.TextFrame.Paragraphs)
{
    //Add the bullets
    para.BulletType = TextBulletType.Numbered;
    para.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod;
}
```

---

# Spire.Presentation C# Text Styling
## Change text style in PowerPoint presentation
```csharp
// Get the shape and its paragraphs
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];
ParagraphCollection paras = shape.TextFrame.Paragraphs;

// Set the style for the text content in the first paragraph
foreach (TextRange tr in paras[0].TextRanges)
{
    tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    tr.Fill.SolidColor.Color = Color.ForestGreen;
    tr.LatinFont = new TextFont("Lucida Sans Unicode");
    tr.FontHeight = 14;
}

// Set the style for the text content in the third paragraph
foreach (TextRange tr in paras[2].TextRanges)
{
    tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    tr.Fill.SolidColor.Color = Color.CornflowerBlue;
    tr.LatinFont = new TextFont("Calibri");
    tr.FontHeight = 16;
    tr.TextUnderlineType = TextUnderlineType.Dashed;
}
```

---

# Spire.Presentation C# Copy Paragraph
## Copy text paragraph from one PowerPoint presentation to another
```csharp
//Load the source file
Presentation ppt1 = new Presentation();
ppt1.LoadFromFile(sourcePath);

//Load the target file
Presentation ppt2 = new Presentation();
ppt2.LoadFromFile(targetPath);

//Get the text from the first shape on the first slide
IShape sourceshp = ppt1.Slides[0].Shapes[0];
string text = ((IAutoShape)sourceshp).TextFrame.Text;

//Get the first shape on the first slide from the target file
IShape destshp = ppt2.Slides[0].Shapes[0];

//Add the text to the target file
((IAutoShape)destshp).TextFrame.Text += "\n\n" + text;

//Save the document
ppt2.SaveToFile(resultPath, FileFormat.Pptx2013);
```

---

# Spire.Presentation C# Custom Bullets
## Customize bullet numbers in PowerPoint presentation
```csharp
//Access the first placeholder in the slide and typecasting it as AutoShape
ITextFrameProperties tf1 = ((IAutoShape)slide.Shapes[1]).TextFrame;

//Access the first Paragraph and set bullet style
TextParagraph para= tf1.Paragraphs[0];
para.Depth = 0;
para.BulletType = TextBulletType.Numbered;
para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod;
para.BulletNumber = 2;

 //Access the second Paragraph and set bullet style
 para = tf1.Paragraphs[1];
 para.Depth = 0;
 para.BulletType = TextBulletType.Numbered;
 para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod;
 para.BulletNumber = 4;

 //Access the third Paragraph and set bullet style
 para = tf1.Paragraphs[2];
 para.Depth = 0;
 para.BulletType = TextBulletType.Numbered;
 para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod;
 para.BulletNumber = 6;

 //Access the fourth Paragraph and set bullet style
 para = tf1.Paragraphs[3];
 para.Depth = 0;
 para.BulletType = TextBulletType.Numbered;
 para.BulletStyle = NumberedBulletStyle.BulletArabicPeriod;
 para.BulletNumber = 7;
```

---

# spire.presentation csharp edit prompt text
## edit prompt text in powerpoint slides
```csharp
// Iterate through the slide
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape.Placeholder != null && shape is IAutoShape)
    {
        string text = "";
        // Set the text of the title
        if (shape.Placeholder.Type == PlaceholderType.CenteredTitle)
        {
            text = "custom title create by Spire";
        }
        // Set text of the subtitle.
        else if (shape.Placeholder.Type == PlaceholderType.Subtitle)
        {
            text = "custom subtitle create by Spire";
        }

        (shape as IAutoShape).TextFrame.Text = text;
    }
}
```

---

# spire.presentation csharp text extraction
## extract text from powerpoint presentation
```csharp
// Create a StringBuilder to store extracted text
StringBuilder sb = new StringBuilder();

// Iterate through each slide in the presentation
foreach (ISlide slide in presentation.Slides)
{
    // Iterate through each shape in the slide
    foreach (IShape shape in slide.Shapes)
    {
        // Check if the shape is an IAutoShape
        if (shape is IAutoShape)
        {
            // Extract text from each paragraph in the shape's text frame
            foreach (TextParagraph tp in (shape as IAutoShape).TextFrame.Paragraphs)
            {
                sb.Append(tp.Text + Environment.NewLine);
            }
        }
    }
}
```

---

# spire.presentation csharp text format
## get default text format from presentation shape
```csharp
// Create Presentation object and load the file
Presentation presentation = new Presentation();
presentation.LoadFromFile(inputFile);

// Get the first shape of the first slide
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

// Get the display format of the text in shape
DefaultTextRangeProperties format = shape.TextFrame.Paragraphs[0].TextRanges[0].DisPlayFormat;

// Determine whether the format is bold or italic
bool isBold = format.IsBold;
bool isItalic = format.IsItalic;

// Dispose
presentation.Dispose();
```

---

# spire.presentation csharp textframe
## get text frame effective data from presentation shape
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
//Get a shape 
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

ITextFrameProperties textFrameFormat = shape.TextFrame;
StringBuilder str = new StringBuilder();
str.AppendLine("Anchoring type: " + textFrameFormat.AnchoringType);
str.AppendLine("Autofit type: " + textFrameFormat.AutofitType);
str.AppendLine("Text vertical type: " + textFrameFormat.VerticalTextType);
str.AppendLine("Margins");
str.AppendLine("   Left: " + textFrameFormat.MarginLeft);
str.AppendLine("   Top: " + textFrameFormat.MarginTop);
str.AppendLine("   Right: " + textFrameFormat.MarginRight);
str.AppendLine("   Bottom: " + textFrameFormat.MarginBottom);
```

---

# spire.presentation text style
## get text style effective data from powerpoint
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Az1.pptx");
//Get the first slide
ISlide slide = presentation.Slides[0];
//Get a shape 
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

StringBuilder str = new StringBuilder();
for (int p = 0; p < shape.TextFrame.Paragraphs.Count; p++)
{   
    var paragraph = shape.TextFrame.Paragraphs[p];
    str.AppendLine("Text style for Paragraph " + p + " :");
    //Get the paragraph style
    str.AppendLine(" Indent: " + paragraph.Indent);
    str.AppendLine(" Alignment: " + paragraph.Alignment);
    str.AppendLine(" Font alignment: " + paragraph.FontAlignment);
    str.AppendLine(" Hanging punctuation: " + paragraph.HangingPunctuation);
    str.AppendLine(" Line spacing: " + paragraph.LineSpacing);
    str.AppendLine(" Space before: " + paragraph.SpaceBefore);
    str.AppendLine(" Space after: " + paragraph.SpaceAfter.ToString());
    str.AppendLine();
    for (int r = 0; r < paragraph.TextRanges.Count; r++)
    {                
        var textRange = paragraph.TextRanges[r];
        str.AppendLine("  Text style for Paragraph " + p + " TextRange " + r + " :");
        //Get the text range style
        str.AppendLine("    Font height: " + textRange.FontHeight);
        str.AppendLine("    Language: " + textRange.Language);
        str.AppendLine("    Font: " + textRange.LatinFont.FontName);
        str.AppendLine();
    }
}

string result = "GetTextStyleEffectiveData_result.txt";
File.WriteAllText(result, str.ToString());
```

---

# Spire.Presentation Text Highlighting
## Highlight specified text in PowerPoint presentation
```csharp
// Get the specified shape
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[1];

TextHighLightingOptions options = new TextHighLightingOptions();
options.WholeWordsOnly = true;
options.CaseSensitive = true;

shape.TextFrame.HighLightText("Spire", Color.Yellow, options);
```

---

# Spire.Presentation Paragraph Indentation
## Set paragraph indentation and spacing in PowerPoint slides
```csharp
// Get shape and paragraphs from the first slide
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];
ParagraphCollection paras = shape.TextFrame.Paragraphs;

// Set the paragraph style for first paragraph
paras[0].Indent = 20;
paras[0].LeftMargin = 10;
paras[0].SpaceAfter = 10;

// Set the paragraph style of the third paragraph 
paras[2].Indent = -100;
paras[2].LeftMargin = 40;
paras[2].SpaceBefore = 0;
paras[2].SpaceAfter = 0;
```

---

# spire.presentation csharp html
## insert html with image into presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();
ShapeList shapes = ppt.Slides[0].Shapes;

shapes.AddFromHtml("<html><div><p>First paragraph</p><p><img src='..\\..\\..\\..\\..\\..\\Data\\Logo.png'/></p><p>Second paragraph </p></html>");
```

---

# spire.presentation csharp line spacing
## set paragraph line spacing in powerpoint slides
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];
//Add a shape 
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, presentation.SlideSize.Size.Width-100,300));
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.TextFrame.Paragraphs.Clear();

//Add text
shape.AppendTextFrame("Spire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers to"
+"create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform."
+"From Spire.Presentation v 3.7.5, Spire.Presentation starts to support .NET Core, .NET standard.");
//Set font and color of text
TextRange textRange = shape.TextFrame.TextRange;
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.BlueViolet;
textRange.FontHeight =20;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");

//Set properties of paragraph
shape.TextFrame.Paragraphs[0].SpaceBefore = 100;
shape.TextFrame.Paragraphs[0].SpaceAfter = 100;
shape.TextFrame.Paragraphs[0].LineSpacing = 150;
```

---

# spire.presentation csharp text formatting
## mix font styles in presentation text
```csharp
//Get the second shape of the first slide
IAutoShape shape = ppt.Slides[0].Shapes[1] as IAutoShape;
//Get the text from the shape 
string originalText = shape.TextFrame.Text;

//Split the string by specified words and return substrings to a string array
string[] splitArray = originalText.Split(new string[] { "bold", "red", "underlined", "bigger font size" }, StringSplitOptions.None);

//Remove the paragraph from TextRange
TextParagraph tp = shape.TextFrame.TextRange.Paragraph;
tp.TextRanges.Clear();

//Append normal text that is in front of 'bold' to the paragraph
TextRange tr = new TextRange(splitArray[0]);
tp.TextRanges.Append(tr);
//Set font style of the text 'bold' as bold
tr = new TextRange("bold");
tr.IsBold = TriState.True;
tp.TextRanges.Append(tr);

//Append normal text that is in front of 'red' to the paragraph
tr = new TextRange(splitArray[1]);
tp.TextRanges.Append(tr);
//Set the color of the text 'red' as red
tr = new TextRange("red");
tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
tr.Format.Fill.SolidColor.Color = Color.Red;
tp.TextRanges.Append(tr);

//Append normal text that is in front of 'underlined' to the paragraph
tr = new TextRange(splitArray[2]);
tp.TextRanges.Append(tr);
//Underline the text 'undelined'
tr = new TextRange("underlined");
tr.TextUnderlineType = TextUnderlineType.Single;
tp.TextRanges.Append(tr);

//Append normal text that is in front of 'bigger font size' to the paragraph
tr = new TextRange(splitArray[3]);
tp.TextRanges.Append(tr);
//Set a large font for the text 'bigger font size'
tr = new TextRange("bigger font size");
tr.FontHeight = 35;
tp.TextRanges.Append(tr);

//Append other normal text
tr = new TextRange(splitArray[4]);
tp.TextRanges.Append(tr);
```

---

# Spire.Presentation C# Text Style Modification
## Modify the style of the first found text in a presentation
```csharp
//Find first "Spire"
string text = "Spire";
TextRange textRange = ppt.Slides[0].FindFirstTextAsRange(text);

//Modify the style
textRange.Fill.FillType = FillFormatType.Solid;
textRange.Fill.SolidColor.Color = Color.Red;
textRange.FontHeight = 28;
textRange.LatinFont = new TextFont("Calibri");
textRange.IsBold = TriState.True;
textRange.IsItalic = TriState.True;
textRange.TextUnderlineType = TextUnderlineType.Double;
textRange.TextStrikethroughType = TextStrikethroughType.Single;
```

---

# spire.presentation csharp bullets
## create multiple level bullets in presentation
```csharp
//Access the first placeholder in the slide and typecasting it as AutoShape
ITextFrameProperties tf1 = ((IAutoShape)slide.Shapes[1]).TextFrame;

//Access the first Paragraph and set bullet style
TextParagraph para= tf1.Paragraphs[0];        
para.BulletType = TextBulletType.Symbol;
para.BulletChar = Convert.ToChar(8226);
para.Depth = 0;

//Access the second Paragraph and set bullet style
para = tf1.Paragraphs[1];
para.BulletType = TextBulletType.Symbol;
para.BulletChar = '-';
para.Depth = 1;

//Access the third Paragraph and set bullet style
para = tf1.Paragraphs[2];
para.BulletType = TextBulletType.Symbol;
para.BulletChar = Convert.ToChar(8226);
para.Depth = 2;

//Access the fourth Paragraph and set bullet style
para = tf1.Paragraphs[3];
para.BulletType = TextBulletType.Symbol;
para.BulletChar = '-';
para.Depth = 3;
```

---

# spire.presentation csharp multiple paragraphs
## create and format multiple paragraphs with text ranges in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Access the first slide
ISlide slide = presentation.Slides[0];

// Add an AutoShape of rectangle type
RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 150, 500, 150);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);

// Access TextFrame of the AutoShape
ITextFrameProperties tf = shape.TextFrame;

// Create Paragraphs and TextRanges with different text formats
TextParagraph para0 = tf.Paragraphs[0];
TextRange textRange1 = new TextRange();
TextRange textRange2 = new TextRange();
para0.TextRanges.Append(textRange1);
para0.TextRanges.Append(textRange2);

TextParagraph para1 = new TextParagraph();
tf.Paragraphs.Append(para1);
TextRange textRange11= new TextRange();
TextRange textRange12 = new TextRange();
TextRange textRange13 = new TextRange();
para1.TextRanges.Append(textRange11);
para1.TextRanges.Append(textRange12);
para1.TextRanges.Append(textRange13);

TextParagraph para2 = new TextParagraph();
tf.Paragraphs.Append(para2);
TextRange textRange21 = new TextRange();
TextRange textRange22 = new TextRange();
TextRange textRange23 = new TextRange();
para2.TextRanges.Append(textRange21);
para2.TextRanges.Append(textRange22);
para2.TextRanges.Append(textRange23);

for (int i = 0; i < 3; i++)
    for (int j = 0; j < 3; j++)
    {
        tf.Paragraphs[i].TextRanges[j].Text = "TextRange " + j.ToString();
        if (j == 0)
        {
            tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid;
            tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.LightBlue;
            tf.Paragraphs[i].TextRanges[j].Format.IsBold = TriState.True;
            tf.Paragraphs[i].TextRanges[j].FontHeight = 15;
        }
        else if (j == 1)
        {
            tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid;
            tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.Blue;
            tf.Paragraphs[i].TextRanges[j].Format.IsItalic = TriState.True;
            tf.Paragraphs[i].TextRanges[j].FontHeight = 18;
        }
    }
```

---

# Spire.Presentation C# Picture Bullet Style
## Customize bullet style with picture in PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the second shape on the first slide
IAutoShape shape = ppt.Slides[0].Shapes[1] as IAutoShape;

//Traverse through the paragraphs in the shape
foreach (TextParagraph paragraph in shape.TextFrame.Paragraphs)
{
    //Set the bullet style of paragraph as picture
    paragraph.BulletType = TextBulletType.Picture;
    //Load a picture
    Image bulletPicture = Image.FromFile("bullet_image.png");
    //Add the picture as the bullet style of paragraph
    paragraph.BulletPicture.EmbedImage = ppt.Images.Append(bulletPicture);
}
```

---

# spire.presentation csharp remove textbox
## remove text box shapes from a powerpoint slide
```csharp
//Get the first slide
ISlide slide = ppt.Slides[0];
//Traverse all the shapes in slide
for (int i = 0; i < slide.Shapes.Count;i++)
{
    if(slide.Shapes[i].Name.Contains("TextBox"))
    {
    	slide.Shapes.RemoveAt(i);
    	i--;
    }
}
```

---

# Spire.Presentation C# Text Replacement
## Replace and format text in PowerPoint presentation
```csharp
// Create a new object to store the default text range formatting properties.
DefaultTextRangeProperties format = new DefaultTextRangeProperties();

// Set the IsBold property of the text range formatting to true, making the text bold.
format.IsBold = TriState.True;

// Set the FillType property of the text range fill to Solid, indicating a solid fill color.
format.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;

// Set the Color property of the solid fill color to red.
format.Fill.SolidColor.Color = Color.Red;

// Set the FontHeight property of the text range formatting to 25, indicating the font size.
format.FontHeight = 25;

// Replace all occurrences of the text "Spire.Presentation for .NET" with "Spire.PPT" and apply the specified formatting.
ppt.ReplaceAndFormatText("Spire.Presentation for .NET", "Spire.PPT", format);
```

---

# spire.presentation csharp text replacement
## replace text in powerpoint presentation slides
```csharp
// Create dictionary with text to replace
Dictionary<string, string> tagValues = new Dictionary<string, string>();
tagValues.Add("Spire.Presentation for .NET", "Spire.PPT");

// Replace text in the first slide
ReplaceTags(ppt.Slides[0], tagValues);

private void ReplaceTags(ISlide pSlide, Dictionary<string, string> TagValues)
{
    foreach (IShape curShape in pSlide.Shapes)
    {
        if (curShape is IAutoShape)
        {
            foreach (TextParagraph tp in (curShape as IAutoShape).TextFrame.Paragraphs)
            {
                foreach (var curKey in TagValues.Keys)
                {
                    if (tp.Text.Contains(curKey))
                    {
                        tp.Text = tp.Text.Replace(curKey, TagValues[curKey]);
                    }
                }
            }
        }
    }
}
```

---

# spire.presentation text replacement
## replace text in powerpoint slides while retaining style
```csharp
// Replace first occurrence of text in slide 0
presentation.Slides[0].ReplaceFirstText("use", "test", true);

// Replace all occurrences of text in slide 1
presentation.Slides[1].ReplaceAllText("Spire", "new spire", true);
```

---

# Spire.Presentation C# Text Replacement
## Replace text in presentation using regular expressions
```csharp
//Create Presentation
Presentation presentation = new Presentation();

//Regex for all words
Regex regex = new Regex(@"\d+.\d+|\w+");

//New string value
string newvalue = "This is the test!";

//Loop and replace
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        shape.ReplaceTextWithRegex(regex, newvalue);
    }
}
```

---

# Spire.Presentation C# Text Rotation
## Rotate text in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];
//Get a shape 
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

shape.TextFrame.VerticalTextType = VerticalTextType.Vertical270;
```

---

# spire.presentation csharp text effects
## set 3D effect for text in presentation
```csharp
//Append a new shape to slide and set the line color and fill type
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(30, 40, 650, 200));
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = FillFormatType.None;

//Add text to the shape
shape.AppendTextFrame("This demo shows how to add 3D effect text to Presentation slide");

//Set the color of text in shape
TextRange textRange = shape.TextFrame.TextRange;
textRange.Fill.FillType = FillFormatType.Solid;
textRange.Fill.SolidColor.Color = Color.LightBlue;

//Set the Font of text in shape
textRange.FontHeight = 40;
textRange.LatinFont = new TextFont("Gulim");

//Set 3D effect for text
shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = PresetMaterialType.Matte;
shape.TextFrame.TextThreeD.LightRig.PresetType = PresetLightRigType.Sunrise;
shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = BevelPresetType.Circle;
shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = Color.Green;
shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3;
```

---

# spire.presentation csharp textframe
## set anchor of text frame in presentation
```csharp
//Get a shape 
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;
shape.TextFrame.AnchoringType = TextAnchorType.Bottom;
```

---

# spire.presentation csharp textframe
## set column count of text frame in presentation
```csharp
//Get the first shape in first slide and set column count of text for it.
IAutoShape shape1 = (IAutoShape)ppt.Slides[0].Shapes[0];
shape1.TextFrame.ColumnCount = 3;

//Get the second shape in second slide and set column count of text for it.
IAutoShape shape2 = (IAutoShape)ppt.Slides[1].Shapes[0];
shape2.TextFrame.ColumnCount = 2;
```

---

# spire.presentation csharp column spacing
## Set column spacing in PowerPoint presentation
```csharp
// Create a new PPT
Presentation presentation = new Presentation();

// Append a shape in the first slide
ISlide slide = presentation.Slides[0];
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 70, 600, 400));
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
shape.Fill.FillType = FillFormatType.None;

// Set column and column spacing
shape.TextFrame.ColumnCount = 2;
shape.TextFrame.ColumnSpacing = 20.50f;
// Append text
shape.TextFrame.Text = "\r\nSpire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers to create, read, write, modify, convert and Print PowerPoint documents on any .NET platform (Target .NET Framework, .NET Core, .NET Standard, .NET 5.0, .NET 6.0, Xamarin & Mono Android). As an independent PowerPoint .NET API, Spire.Presentation for .NET doesn't need Microsoft PowerPoint to be installed on machines.\r\n\r\n\r\nSpire.Presentation for .NET supports PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also supports exporting presentation slides to images (PNG, JPG, TIFF, EMF, SVG), PDF, XPS, HTML format etc.";
foreach (TextParagraph paragraph in shape.TextFrame.Paragraphs)
{
    foreach (TextRange textRange in paragraph.TextRanges)
    {
        // Set font for text
        textRange.Fill.FillType = FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = Color.Black;
        textRange.FontHeight = 16;
        textRange.LatinFont = new TextFont("Open Sans");
    }
}
```

---

# Spire.Presentation C# Custom Fonts
## Setting custom fonts in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add a shape with text
RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 80, 500, 150);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.AppendTextFrame("Hello World!");

//Set the custom font folder
presentation.SetCustomFontsFolder(@"E:\customFonts\");

//Set the font and fill style of the text
TextRange textRange = shape.TextFrame.TextRange;
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
textRange.FontHeight = 66;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");
```

---

# spire.presentation csharp paragraph font
## Set paragraph font properties in PowerPoint presentation
```csharp
//Access the first and second placeholder in the slide and typecasting it as AutoShape
ITextFrameProperties tf1 = ((IAutoShape)slide.Shapes[0]).TextFrame;
ITextFrameProperties tf2 = ((IAutoShape)slide.Shapes[1]).TextFrame;

// Access the first Paragraph
TextParagraph para1 = tf1.Paragraphs[0];
TextParagraph para2 = tf2.Paragraphs[0];

//Justify the paragraph
para2.Alignment = TextAlignmentType.Justify;

//Access the first text range
TextRange textRange1 = para1.FirstTextRange;
TextRange textRange2 = para2.FirstTextRange;

//Define new fonts
TextFont fd1 = new TextFont("Elephant");
TextFont fd2 = new TextFont("Castellar");
 
// Assign new fonts to text range
textRange1.LatinFont = fd1;
textRange2.LatinFont = fd2;

// Set font to Bold
textRange1.Format.IsBold = TriState.True;
textRange2.Format.IsBold = TriState.False;

// Set font to Italic
textRange1.Format.IsItalic = TriState.False;
textRange2.Format.IsItalic = TriState.True;

// Set font color
textRange1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange1.Fill.SolidColor.Color = Color.Purple;
textRange2.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange2.Fill.SolidColor.Color = Color.Peru;
```

---

# spire.presentation right-to-left columns
## Set text frame to right-to-left column mode in PowerPoint presentation
```csharp
// Get the second shape
IAutoShape shape = ppt.Slides[0].Shapes[1] as IAutoShape;
// Set columns style to right-to-left
shape.TextFrame.RightToLeftColumns = true;
```

---

# Spire.Presentation C# Text Shadow Effect
## Apply outer shadow effect to text in presentation slides
```csharp
//Add outer shadow and set all necessary parameters
OuterShadowEffect Shadow = new OuterShadowEffect();

Shadow.BlurRadius = 0;
Shadow.Direction = 50;
Shadow.Distance = 10;
Shadow.ColorFormat.Color = Color.LightBlue;

//Apply the outer shadow effect to text in a shape
shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow;
```

---

# spire.presentation csharp text direction
## set text direction in presentation shapes
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a shape with text to the first slide
IAutoShape textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 70, 100, 400));
textboxShape.ShapeStyle.LineColor.Color = Color.Transparent;
textboxShape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textboxShape.Fill.SolidColor.Color = Color.LightBlue;
textboxShape.TextFrame.Text = "You Are Welcome Here";
//Set the text direction to vertical
textboxShape.TextFrame.VerticalTextType = VerticalTextType.Vertical;

//Append another shape with text to the slide
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(350, 70, 100, 400));
textboxShape.ShapeStyle.LineColor.Color = Color.Transparent;
textboxShape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textboxShape.Fill.SolidColor.Color = Color.LightGray;
//Append some asian characters
textboxShape.TextFrame.Text = "欢迎光临";
//Set the VerticalTextType as EastAsianVertical to aviod rotating text 90 degrees
textboxShape.TextFrame.VerticalTextType = VerticalTextType.EastAsianVertical;
```

---

# Spire.Presentation C# Text Font Properties
## Set text font properties in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add a new shape to the PPT document
RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 250, 80, 500, 150);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);

shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

//Add text to the shape
shape.AppendTextFrame("Welcome to use Spire.Presentation");

TextRange textRange = shape.TextFrame.TextRange;
//Set the font
textRange.LatinFont = new TextFont("Times New Roman");
//Set bold property of the font
textRange.IsBold = TriState.True;

//Set italic property of the font
textRange.IsItalic = TriState.True;

//Set underline property of the font
textRange.TextUnderlineType = TextUnderlineType.Single;

//Set the height of the font
textRange.FontHeight = 50;

//Set the color of the font
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
```

---

# Spire.Presentation C# Text Margins
## Set text margins for shapes in PowerPoint presentations
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a new shape
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, 450, 150));

//Set text in the shape
shape.TextFrame.Text = "Using Spire.Presentation, developers will find an easy and effective method to create, read, write, modify, convert and print PowerPoint files on any .Net platform. It's worthwhile for you to try this amazing product.";

//Set the margins for the text frame
shape.TextFrame.MarginTop = 10;
shape.TextFrame.MarginBottom = 35;
shape.TextFrame.MarginLeft = 15;
shape.TextFrame.MarginRight = 30;
```

---

# Spire.Presentation C# Text Transparency
## Set text transparency with different alpha values in a presentation
```csharp
//Add a shape
IAutoShape textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 300, 120));
textboxShape.ShapeStyle.LineColor.Color = Color.Transparent;
textboxShape.Fill.FillType = FillFormatType.None;

//Remove default blank paragraphs
textboxShape.TextFrame.Paragraphs.Clear();

//Add three paragraphs, apply color with different alpha values to text
int alpha = 55;
for (int i = 0; i < 3; i++)
{
    textboxShape.TextFrame.Paragraphs.Append(new TextParagraph());
    textboxShape.TextFrame.Paragraphs[i].TextRanges.Append(new TextRange("Text Transparency"));
    textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.FillType = FillFormatType.Solid;
    textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(alpha, Color.Purple);
    alpha += 100;
}
```

---

# spire.presentation superscript subscript
## create superscript and subscript text in presentation
```csharp
//Add a shape 
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 100, 200, 50));
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.TextFrame.Paragraphs.Clear();

shape.AppendTextFrame("Test");
TextRange tr = new TextRange("superscript");
shape.TextFrame.Paragraphs[0].TextRanges.Append(tr);

//Set superscript text
shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = 30;

TextRange textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.Black;
textRange.FontHeight = 20;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");

textRange = shape.TextFrame.Paragraphs[0].TextRanges[1];
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");

shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 150, 200, 50));
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.TextFrame.Paragraphs.Clear();

shape.AppendTextFrame("Test");
tr = new TextRange("subscript");
shape.TextFrame.Paragraphs[0].TextRanges.Append(tr);

//Set subscript text
shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = -25;

textRange = shape.TextFrame.Paragraphs[0].TextRanges[0];
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.Black;
textRange.FontHeight = 20;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");

textRange = shape.TextFrame.Paragraphs[0].TextRanges[1];
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = System.Drawing.Color.CadetBlue;
textRange.LatinFont = new TextFont("Lucida Sans Unicode");
```

---

# spire.presentation csharp slides
## add different layout slides to presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Remove the default slide
presentation.Slides.RemoveAt(0);

//Loop through slide layouts
foreach (SlideLayoutType type in Enum.GetValues(typeof(SlideLayoutType)))
{
    //Append slide by specifing slide layout
    presentation.Slides.Append(type);
}
```

---

# Spire.Presentation C# Add Image to Master
## Demonstrates how to add an image to a master slide in a PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the master collection
IMasterSlide master = presentation.Masters[0];

//Append image to slide master
String image = @"..\..\..\..\..\..\Data\Logo.png";
RectangleF rff = new RectangleF(40, 40, 90, 90);
IEmbedImage pic = master.Shapes.AppendEmbedImage(ShapeType.Rectangle, image, rff);
pic.Line.FillFormat.FillType = FillFormatType.None;

//Add new slide to presentation
presentation.Slides.Append();
```

---

# spire.presentation csharp slide management
## add slides using master layout
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get Master layouts
ILayout iLayout = presentation.Masters[0].Layouts[0];

//Append new slide
presentation.Slides.Append(iLayout);

//Insert new slide
presentation.Slides.Insert(1, iLayout);
```

---

# Spire.Presentation C# Slide Management
## Append slides with master layout
```csharp
//Get the master
IMasterSlide master = presentation.Masters[0];

//Get master layout slides
IMasterLayouts masterLayouts = master.Layouts;
ActiveSlide layoutSlide = masterLayouts[1] as ActiveSlide;

//Append a rectangle to the layout slide
IAutoShape shape = layoutSlide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(10, 50, 100, 80));

//Add a text into the shape and set the style
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.AppendTextFrame("Layout slide 1");
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Black");
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.CadetBlue;

//Append new slide with master layout
presentation.Slides.Append(presentation.Slides[0], master.Layouts[1]);

//Another way to append new slide with master layout
presentation.Slides.Insert(2, presentation.Slides[1], master.Layouts[1]);
```

---

# spire.presentation csharp slide master
## apply slide master to presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide master from the presentation
IMasterSlide masterSlide = ppt.Masters[0];

//Customize the background of the slide master
RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
masterSlide.SlideBackground.Fill.FillType = FillFormatType.Picture;
IEmbedImage image = masterSlide.Shapes.AppendEmbedImage(ShapeType.Rectangle, backgroundPic, rect);
masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image as IImageData;

//Change the color scheme
masterSlide.Theme.ColorScheme.Accent1.Color = Color.Red;
masterSlide.Theme.ColorScheme.Accent2.Color = Color.RosyBrown;
masterSlide.Theme.ColorScheme.Accent3.Color = Color.Ivory;
masterSlide.Theme.ColorScheme.Accent4.Color = Color.Lavender;
masterSlide.Theme.ColorScheme.Accent5.Color = Color.Black;
```

---

# spire.presentation slide transitions
## Set different transition types and timings for presentation slides
```csharp
//Set the first slide transition as circle
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle;

// Set the transition time of 3 seconds
presentation.Slides[0].SlideShowTransition.AdvanceOnClick = true;
presentation.Slides[0].SlideShowTransition.AdvanceAfterTime = 3000;

//Set the second slide transition as comb and set the speed 
presentation.Slides[1].SlideShowTransition.Type = TransitionType.Comb;
presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow;

// Set the transition time of 5 seconds
presentation.Slides[1].SlideShowTransition.AdvanceOnClick = true;
presentation.Slides[1].SlideShowTransition.AdvanceAfterTime = 5000;

// Set the third slide transition as zoom
presentation.Slides[2].SlideShowTransition.Type = TransitionType.Zoom;

// Set the transition time of 7 seconds
presentation.Slides[2].SlideShowTransition.AdvanceOnClick = true;
presentation.Slides[2].SlideShowTransition.AdvanceAfterTime = 7000;
```

---

# Spire.Presentation C# Slide Layout
## Change the layout of a slide in a PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile("ChangeSlideLayout.pptx");

//Change the layout of slide
presentation.Slides[1].Layout = presentation.Masters[0].Layouts[4];
```

---

# spire.presentation csharp slide manipulation
## change slide position in presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Move the first slide to the second slide position
ISlide slide = presentation.Slides[0];
slide.SlideNumber = 2;
```

---

# Spire.Presentation C# Clone Slides
## Clone slides from one PowerPoint presentation to another
```csharp
// Create source and destination presentations
Presentation sourcePPT = new Presentation();
Presentation destPPT = new Presentation();

// Loop through all slides of source document
foreach (ISlide slide in sourcePPT.Slides)
{
    // Append the slide at the end of destination document
    destPPT.Slides.Append(slide);
}
```

---

# spire.presentation csharp clone masters
## clone master slides from one presentation to another
```csharp
//Add masters from PPT1 to PPT2
foreach (IMasterSlide masterSlide in presentation1.Masters)
{
    presentation2.Masters.AppendSlide(masterSlide);
}
```

---

# spire.presentation csharp slide manipulation
## clone slide at the end of presentation
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];

//Append the slide at the end of the document
presentation.Slides.Append(slide);
```

---

# Spire.Presentation C# Slide Cloning
## Clone a slide from one PowerPoint presentation to another
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile("source1.pptx");

//Load the another document and choose the first slide to be cloned
Presentation ppt1 = new Presentation();
ppt1.LoadFromFile("source2.pptx");
ISlide slide1 = ppt1.Slides[0];

//Insert the slide to the specified index in the source presentation
int index = 1;
presentation.Slides.Insert(index, slide1); 

//Save the document
presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);
```

---

# spire.presentation csharp slide cloning
## clone slide within presentation
```csharp
//Get a list of slides and choose the first slide to be cloned
ISlide slide = ppt.Slides[0];

//Insert the desired slide to the specified index in the same presentation
int index = 1;
ppt.Slides.Insert(index, slide);
```

---

# spire.presentation csharp slide creation
## create and format PowerPoint slides with shapes and text
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Add new slide
presentation.Slides.Append();

//Add title
RectangleF rec_title = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 70, 400, 50);
IAutoShape shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title);
shape_title.ShapeStyle.LineColor.Color = Color.White;
shape_title.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
TextParagraph para_title = new TextParagraph();
para_title.Text = "E-iceblue";
para_title.Alignment = TextAlignmentType.Center;
para_title.TextRanges[0].LatinFont = new TextFont("Myriad Pro Light");
para_title.TextRanges[0].FontHeight = 36;
para_title.TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
para_title.TextRanges[0].Fill.SolidColor.Color = Color.Black;
shape_title.TextFrame.Paragraphs.Append(para_title);

//Append new shape
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 150, 600, 280));
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.Line.FillType = FillFormatType.None;
//Add text to shape
shape.AppendTextFrame("Welcome to use Spire.Presentation for .NET.");

//Add new paragraph
TextParagraph pare = new TextParagraph();
pare.Text = "";
shape.TextFrame.Paragraphs.Append(pare);

//Add new paragraph
pare = new TextParagraph();
pare.Text = "Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.";
shape.TextFrame.Paragraphs.Append(pare);

//Set the Font
foreach (TextParagraph para in shape.TextFrame.Paragraphs)
{
    para.TextRanges[0].LatinFont = new TextFont("Myriad Pro");
    para.TextRanges[0].FontHeight = 24;
    para.TextRanges[0].Fill.FillType = FillFormatType.Solid;
    para.TextRanges[0].Fill.SolidColor.Color = Color.Black;
    para.Alignment = TextAlignmentType.Left;
}

//Append new shape - SixPointedStar
shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.SixPointedStar, new RectangleF(100, 100, 100, 100));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.Orange;
shape.ShapeStyle.LineColor.Color = Color.White;

//Append new shape
shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 250, 600, 50));
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;

//Add text to shape
shape.AppendTextFrame("This is newly added Slide.");

//Set the Font
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Myriad Pro");
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24;
shape.TextFrame.Paragraphs[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].Fill.SolidColor.Color = Color.Black;
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
shape.TextFrame.Paragraphs[0].Indent = 35;
```

---

# spire.presentation csharp slide master
## create slide master and apply to slides
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

ppt.SlideSize.Type = SlideSizeType.Screen16x9;

//Add slides
for (int i = 0; i < 4; i++)
{
    ppt.Slides.Append();
}

//Get the first default slide master
IMasterSlide first_master = ppt.Masters[0];

//Append another slide master
ppt.Masters.AppendSlide(first_master);
IMasterSlide second_master = ppt.Masters[1];

//Set different background image for the two slide masters
//The first slide master
RectangleF rect = new RectangleF(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
first_master.SlideBackground.Fill.FillType = FillFormatType.Picture;
IEmbedImage image1 = first_master.Shapes.AppendEmbedImage(ShapeType.Rectangle, "image_path", rect);
first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image1 as IImageData;
//The second slide master
second_master.SlideBackground.Fill.FillType = FillFormatType.Picture;
IEmbedImage image2 = second_master.Shapes.AppendEmbedImage(ShapeType.Rectangle, "image_path", rect);
second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image2 as IImageData;

//Apply the first master with layout to the first slide
ppt.Slides[0].Layout = first_master.Layouts[1];

//Apply the second master with layout to other slides
for (int i = 1; i < ppt.Slides.Count; i++)
{
    ppt.Slides[i].Layout = second_master.Layouts[8];
}
```

---

# spire.presentation csharp theme detection
## detect used themes in presentation slides
```csharp
// Create an instance of presentation document
Presentation ppt = new Presentation();

StringBuilder sb = new StringBuilder();
string themeName = null;
sb.AppendLine("This is the name list of the used theme below.");
// Get the theme name of each slide in the document
foreach (ISlide slide in ppt.Slides)
{
    themeName = slide.Theme.Name;
    sb.AppendLine(themeName);
}
```

---

# spire.presentation csharp slide transition
## disable advance after time setting for slide transition
```csharp
// Create a Presentation object
Presentation ppt = new Presentation();

// Load the PPT file
ppt.LoadFromFile("input.pptx");

// Get the first slide and disable the selected advance after time setting
ppt.Slides[0].SlideShowTransition.SelectedAdvanceAfterTime = false;
```

---

# spire.presentation get slide
## retrieve slides by index or ID
```csharp
//Get slide by index 0
ISlide slide1 = presentation.Slides[0];
//Append a shape in the slide
IAutoShape shape1 = slide1.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 200, 100));
//Add text in the shape
shape1.TextFrame.Text = "Get slide by index";

//Get slide by slide ID
ISlide slide2 = presentation.FindSlide((int)presentation.Slides[1].SlideID);
//Append a shape in the slide
IAutoShape shape2 = slide2.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 200, 100));
//Add text in the shape
shape2.TextFrame.Text = "Get slide by slide id";
```

---

# spire.presentation csharp slide layout
## get slide layout names from presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

StringBuilder builder = new StringBuilder();

//Loop through the slides of PPT document
for (int i = 0; i < presentation.Slides.Count; i++)
{
    //Get the name of slide layout
    string name = presentation.Slides[i].Layout.Name;
    builder.AppendLine(String.Format("The name of slide {0} layout is: {1}", i, name));
}
```

---

# spire.presentation csharp get slide text
## extract text from all slides in a powerpoint presentation
```csharp
// Create a PPT document and load file
Presentation ppt = new Presentation();
ppt.LoadFromFile(@"..\..\..\..\..\..\Data\GetSlideText.pptx");

// Iterate through each slide and extract text
foreach (ISlide slide in ppt.Slides)
{
    ArrayList arrayList = slide.GetAllTextFrame();
    foreach (String text in arrayList)
    {
        MessageBox.Show(text);
    }
}
```

---

# Spire.Presentation C# Hide Slide
## Hide a specific slide in a PowerPoint presentation
```csharp
//Create a PPT document and load PPT file from disk
Presentation ppt = new Presentation();
ppt.LoadFromFile("presentation.pptx");

//Hide the second slide
ppt.Slides[1].Hidden = true;
```

---

# Spire.Presentation C# Placeholder
## Insert a placeholder in a master slide
```csharp
// Create a Presentation object
Presentation presentation = new Presentation();

// Insert placeholder
presentation.Masters[0].Layouts[0].InsertPlaceholder(InsertPlaceholderType.Text, new RectangleF(20, 30, 400, 400));
```

---

# spire.presentation csharp image conversion
## convert PowerPoint master layouts to images
```csharp
// Create a PPT document and load the file
Presentation ppt = new Presentation();
ppt.LoadFromFile(@"..\..\..\..\..\..\Data\CloneMaster2.pptx");

// Iterate the masters
for (int i = 0; i < ppt.Masters[0].Layouts.Count; i++)
{
    // Save layouts as images
    Image image = ppt.Masters[0].Layouts[i].SaveAsImage();
    String fileName = String.Format("{0}.png", i);
    image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
}

ppt.Dispose();
```

---

# spire.presentation csharp slides
## merge selected slides from different presentations
```csharp
// Create an instance of presentation document
Presentation ppt = new Presentation();

// Remove the first slide
ppt.Slides.RemoveAt(0);

// Load two PPT files
Presentation ppt1 = new Presentation(@"..\..\..\..\..\..\Data\InputTemplate.pptx", FileFormat.Pptx2013);
Presentation ppt2 = new Presentation(@"..\..\..\..\..\..\Data\TextTemplate.pptx", FileFormat.Pptx2013);

// Append all slides in ppt1 to ppt
for (int i = 0; i < ppt1.Slides.Count; i++)
{
    ppt.Slides.Append(ppt1.Slides[i]);
}

// Append the second slide in ppt2 to ppt
ppt.Slides.Append(ppt2.Slides[1]);
```

---

# Spire.Presentation C# Slide Removal
## Demonstrates how to remove slides from a PowerPoint presentation using Spire.Presentation library
```csharp
Presentation presentation = new Presentation();

//Remove slide by index
presentation.Slides.RemoveAt(0);

//Remove slide by its reference
ISlide slide = presentation.Slides[1];
presentation.Slides.Remove(slide);
```

---

# spire.presentation remove unused layout
## Remove unused layout masters from PowerPoint presentation
```csharp
//Create an array list
List<IActiveSlide> list = new List<IActiveSlide>();
for (int i = 0; i < ppt.Slides.Count; i++)
{
    //Get the layout used by slide
    IActiveSlide layout = (IActiveSlide)ppt.Slides[i].Layout;
    list.Add(layout);
}

//Loop through masters and layouts
for (int i = 0; i < ppt.Masters.Count; i++)
{
    IMasterLayouts masterlayouts = ppt.Masters[i].Layouts;
    for (int j = masterlayouts.Count - 1; j >= 0; j--)
    {
        if (!list.Contains((IActiveSlide)masterlayouts[j]))
        {
            //Remove unused layout
            masterlayouts.RemoveMasterLayout(j);
        }
    }
}
```

---

# Spire.Presentation C# Slide Transition
## Set advance after time for slides in a PowerPoint presentation
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//Traverse all slides
for (int i = 0; i < ppt.Slides.Count; i++)
{
    ppt.Slides[i].SlideShowTransition.AdvanceOnClick = true;

    //Set the advance after time to 5000 milliseconds
    ppt.Slides[i].SlideShowTransition.AdvanceAfterTime = 5000;
}
```

---

# spire.presentation csharp slide layout
## set slide layout and add content
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Remove the first slide
ppt.Slides.RemoveAt(0);

//Append a slide and set the layout for slide
ISlide slide = ppt.Slides.Append(SlideLayoutType.Title);

//Add content for Title and Text
IAutoShape shape = slide.Shapes[0] as IAutoShape;
shape.TextFrame.Text = "Hello Wolrd! –> This is title";

shape = slide.Shapes[1] as IAutoShape;
shape.TextFrame.Text = "E-iceblue Support Team -> This is content";
```

---

# spire.presentation csharp slide numbering
## Set starting number for slides in a presentation
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk.
presentation.LoadFromFile("ChangeSlidePosition.pptx");

//Set 5 as the starting number
presentation.FirstSlideNumber = 5;
```

---

# spire.presentation csharp slide transition
## set slide transition effects
```csharp
// Set effects
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Cut;
((OptionalBlackTransition)presentation.Slides[0].SlideShowTransition.Value).FromBlack = true;
```

---

# spire.presentation csharp slide transitions
## set slide transitions in presentation
```csharp
//Set the first slide transition as push and sound mode
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Push;
presentation.Slides[0].SlideShowTransition.SoundMode = TransitionSoundMode.StartSound;

//Set the second slide transition as circle and set the speed 
presentation.Slides[1].SlideShowTransition.Type = TransitionType.Fade;
presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow;
```

---

# spire.presentation csharp master background
## show master background graphics in presentation slides
```csharp
// Create a Presentation object
Presentation presentation = new Presentation();

// Set whether to show the background graphics of the slide master
presentation.Slides[0].Layout.ShowMasterShapes = true;
```

---

# spire.presentation slide title
## get and set slide titles in PowerPoint presentation
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
//Get the title of the first slide
String slideTitle = slide.Title;

//Set the title of the second slide
presentation.Slides[1].Title = "Second Slide";
```

---

# spire.presentation csharp math equations
## Add and detect math equations in PowerPoint presentation
```csharp
//Create Presentation
Presentation presentation = new Presentation();

//Math code
string latexMathCode = @"x^{2}+\sqrt{x^{2}+1}=2";

//Append a shape
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, new RectangleF(30, 100, 400, 30));
shape.TextFrame.Paragraphs.Clear();

//Add math equation
TextParagraph tp = shape.TextFrame.Paragraphs.AddParagraphFromLatexMathCode(latexMathCode);

//Detect if the slide contains math equation
for (int i = 0; i < presentation.Slides[0].Shapes.Count; i++)
{
    if (presentation.Slides[0].Shapes[i] is IAutoShape)
    {
        bool containMathEquation = (presentation.Slides[0].Shapes[i] as IAutoShape).ContainMathEquation;
        MessageBox.Show("The first slide contains math equations: " + containMathEquation);
    }
}
```

---

# Spire.Presentation C# Animation
## Add exit animation to shapes in presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Add a shape to the slide
IShape starShape = slide.Shapes.AppendShape(ShapeType.FivePointedStar, new RectangleF(250, 100, 200, 200));
starShape.Fill.FillType = FillFormatType.Solid;
starShape.Fill.SolidColor.KnownColor = KnownColors.LightBlue;

//Add random bars effect to the shape
AnimationEffect effect = slide.Timeline.MainSequence.AddEffect(starShape, AnimationEffectType.RandomBars);

//Change effect type from entrance to exit
effect.PresetClassType = TimeNodePresetClassType.Exit;
```

---

# Spire.Presentation C# Line Shape
## Add a line to a PowerPoint slide
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];

//Add a line in the slide
IAutoShape line = slide.Shapes.AppendShape(ShapeType.Line, new RectangleF(50, 100, 300, 0));

//Set color of the line
line.ShapeStyle.LineColor.Color = Color.Red;
```

---

# Spire.Presentation C# Line with Arrow
## Add lines with different arrow types to a presentation slide
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Add a line to the slides and set its color to red
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, new RectangleF(150, 100, 100, 100));
shape.ShapeStyle.LineColor.Color = Color.Red;
//Set the line end type as StealthArrow
shape.Line.LineEndType = LineEndType.StealthArrow;

//Add a line to the slides and use default color
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, new RectangleF(300, 150, 100, 100));
shape.Rotation = -45;
//Set the line end type as TriangleArrowHead
shape.Line.LineEndType = LineEndType.TriangleArrowHead;

//Add a line to the slides and set its color to Green
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, new RectangleF(450, 100, 100, 100));
shape.ShapeStyle.LineColor.Color = Color.Green;
shape.Rotation = 90;
//Set the line begin type as TriangleArrowHead
shape.Line.LineBeginType = LineEndType.StealthArrow;
```

---

# spire.presentation csharp add line
## Add lines with two points to a presentation slide
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Add line with two points
IAutoShape line = slide.Shapes.AppendShape(ShapeType.Line, new PointF(50, 50), new PointF(150, 150));
line.ShapeStyle.LineColor.Color = Color.Red;
line = slide.Shapes.AppendShape(ShapeType.Line, new PointF(150, 150), new PointF(250, 50));
line.ShapeStyle.LineColor.Color = Color.Blue;
```

---

# Spire.Presentation MathML Equation
## Adding MathML equation to PowerPoint slide
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//Set the mathML code
String mathMLCode = "<mml:math xmlns:mml=\"http://www.w3.org/1998/Math/MathML\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">" + "<mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:msqrt><mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:msqrt><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:math>";

//Add a shape
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new Rectangle(30, 100, 400, 30));
shape.TextFrame.Paragraphs.Clear();

//Add the mathml equation paragraph
TextParagraph tp = shape.TextFrame.Paragraphs.AddParagraphFromMathMLCode(mathMLCode);
```

---

# Spire.Presentation C# Round Corner Rectangle
## Create a presentation and add a round corner rectangle with custom styling
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a round corner rectangle and set its radius
IAutoShape shape = ppt.Slides[0].Shapes.AppendRoundRectangle(300, 90, 100, 200, 80);
//Set the color and fill style of shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightBlue;
shape.ShapeStyle.LineColor.Color = Color.SkyBlue;
//Rotate the shape to 90 degree
shape.Rotation = 90;
```

---

# spire.presentation csharp shapes
## add various shapes to powerpoint slide
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Set background Image
RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

//Append new shape - Triangle and set style
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(115, 130, 100, 100));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightGreen;
shape.ShapeStyle.LineColor.Color = Color.White;

//Append new shape - Ellipse
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(290, 130, 150, 100));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightSkyBlue;
shape.ShapeStyle.LineColor.Color = Color.White;

//Append new shape - Heart
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Heart, new RectangleF(470, 130, 130, 100));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.Red;
shape.ShapeStyle.LineColor.Color = Color.LightGray;

//Append new shape - FivePointedStar
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.FivePointedStar, new RectangleF(90, 270, 150, 150));
shape.Fill.FillType = FillFormatType.Gradient;
shape.Fill.SolidColor.Color = Color.Black;
shape.ShapeStyle.LineColor.Color = Color.White;

//Append new shape - Rectangle
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(320, 290, 100, 120));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.Pink;
shape.ShapeStyle.LineColor.Color = Color.LightGray;

//Append new shape - BentUpArrow
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.BentUpArrow, new RectangleF(470, 300, 150, 100));

//Set the color of shape
shape.Fill.FillType = FillFormatType.Gradient;
shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.Olive);
shape.Fill.Gradient.GradientStops.Append(0, KnownColors.PowderBlue);
shape.ShapeStyle.LineColor.Color = Color.White;
```

---

# spire.presentation animation effects
## create slide animations and shape animation effects
```csharp
//Set the animation of slide to Circle
presentation.Slides[0].SlideShowTransition.Type = Spire.Presentation.Drawing.Transition.TransitionType.Circle;

//Append new shape - Triangle
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, new RectangleF(100, 280, 80, 80));

//Set the color of shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.CadetBlue;
shape.ShapeStyle.LineColor.Color = Color.White;

//Set the animation of shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Path4PointStar);

//Append new shape - Rectangle and set animation
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(210, 280, 150, 80));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.CadetBlue;
shape.ShapeStyle.LineColor.Color = Color.White;
shape.AppendTextFrame("Animated Shape");
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel);

//Append new shape - Cloud and set the animation
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Cloud, new RectangleF(390, 280, 80, 80));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
shape.ShapeStyle.LineColor.Color = Color.CadetBlue;
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedZoom);
```

---

# spire.presentation csharp chart animation
## apply animation effect to chart in presentation
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
//Get chart
IShape shape = slide.Shapes[0];
if (shape is IChart)
{ 
    //Apply Wipe animation effect to the chart
    AnimationEffect effect = slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Wipe);
    //Set the BuildType as Series
    effect.GraphicAnimation.BuildType = GraphicBuildType.BuildAsSeries;
}
```

---

# Spire.Presentation C# Animation
## Apply animation effect to a shape in a presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Insert a rectangle in the slide and fill the shape
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 150, 200, 80));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightBlue;
shape.ShapeStyle.LineColor.Color = Color.White;
shape.AppendTextFrame("Animated Shape");

//Apply FadedSwivel animation effect to the shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel);
```

---

# Spire.Presentation C# Animation
## Apply animation on text in PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Add a shape to the slide
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 150, 200, 100));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightBlue;
shape.ShapeStyle.LineColor.Color = Color.White;
shape.AppendTextFrame("This demo shows how to apply animation on text in PPT document.");

//Apply animation to the text in shape
AnimationEffect animation = shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Float);
animation.SetStartEndParagraphs(0, 0);
```

---

# Spire.Presentation C# Shape Arrangement
## Arrange shapes in a PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the specified shape
IShape shape = ppt.Slides[0].Shapes[0];

//Bring the shape forward through SetShapeArrange method
shape.SetShapeArrange(ShapeArrange.BringForward);
```

---

# spire.presentation csharp background
## set background image for presentation slide
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Set background Image
string ImageFile = @"..\..\..\..\..\..\Data\backgroundImg.png";
RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
```

---

# Spire.Presentation C# Shape Copying
## Copy shapes between slides in a presentation
```csharp
// Define the source slide and target slide
ISlide sourceSlide = ppt.Slides[0];
ISlide targetSlide = ppt.Slides[1];

// Copy the first shape from the source slide to the target slide
targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[0]);
```

---

# spire.presentation csharp custom animation
## create custom path animation for shapes in presentation
```csharp
//Create PPT document
Presentation ppt = new Presentation();

//Add shape
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(0, 0, 200, 200));

//Add animation
AnimationEffect effect = ppt.Slides[0].Timeline.MainSequence.
    AddEffect(shape, AnimationEffectType.PathUser);
CommonBehaviorCollection common = effect.CommonBehaviorCollection;
AnimationMotion motion = (AnimationMotion)common[0];
motion.Origin = AnimationMotionOrigin.Layout;
motion.PathEditMode = AnimationMotionPathEditMode.Relative;

//Add motion path
MotionPath moinPath = new MotionPath();
moinPath.Add(MotionCommandPathType.MoveTo, new PointF[] { new PointF(0, 0) }, MotionPathPointsType.CurveAuto, true);
moinPath.Add(MotionCommandPathType.LineTo, new PointF[] { new PointF(0.1f, 0.1f) }, MotionPathPointsType.CurveAuto, true);
moinPath.Add(MotionCommandPathType.LineTo, new PointF[] { new PointF(-0.1f, 0.2f) }, MotionPathPointsType.CurveAuto, true);
moinPath.Add(MotionCommandPathType.End, new PointF[] { }, MotionPathPointsType.CurveStraight, true);
motion.Path = moinPath;
```

---

# spire.presentation animation timing
## manage animation duration and delay time in PowerPoint
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
AnimationEffectCollection animations = slide.Timeline.MainSequence;

//Get duration time of animation
float durationTime = animations[0].Timing.Duration;

//Set new duration time of animation
animations[0].Timing.Duration = 0.8f;

//Get delay time of animation
float delayTime = animations[0].Timing.TriggerDelayTime;

//Set new delay time of animation
animations[0].Timing.TriggerDelayTime = 0.6f;
```

---

# Spire.Presentation C# SVG Embedding
## Embed SVG file into PowerPoint presentation
```csharp
// Define the input SVG file path
string inputFile = @"..\..\..\..\..\..\Data\charthtml.svg";            

// Create Presentation object
Presentation presentation = new Presentation();

// Embed svg in presentation shape
presentation.Slides[0].Shapes.AddFromSVG(inputFile, new RectangleF(40, 40, 200, 200));
```

---

# spire.presentation csharp gradient fill
## fill shape with gradient in presentation
```csharp
//Get the first shape and set the style to be Gradient
IAutoShape GradientShape = ppt.Slides[0].Shapes[0] as IAutoShape;
GradientShape.Fill.FillType = FillFormatType.Gradient;
GradientShape.Fill.Gradient.GradientStops.Append(0, Color.LightSkyBlue);
GradientShape.Fill.Gradient.GradientStops.Append(1, Color.LightGray);
```

---

# spire.presentation csharp pattern fill
## fill shape with pattern in presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];

//Add a rectangle
RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 50, 100, 100, 100);
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect);

//Set the pattern fill format 
shape.Fill.FillType = FillFormatType.Pattern;
shape.Fill.Pattern.PatternType = PatternFillType.Trellis;
shape.Fill.Pattern.BackgroundColor.Color = Color.DarkGray;
shape.Fill.Pattern.ForegroundColor.Color = Color.Yellow;

//Set the fill format of line
shape.Line.FillType = FillFormatType.Solid;
shape.Line.SolidFillColor.Color = Color.Transparent;
```

---

# spire.presentation csharp shape
## fill shape with picture
```csharp
//Get the first shape
IAutoShape shape = ppt.Slides[0].Shapes[0] as IAutoShape;

//Fill the shape with picture
shape.Fill.FillType = FillFormatType.Picture;
shape.Fill.PictureFill.Picture.Url = "image_path";
shape.Fill.PictureFill.FillType = PictureFillType.Stretch;
```

---

# spire.presentation csharp shape fill
## fill shape with solid color in powerpoint
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];

//Add a rectangle
RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 50, 100, 100, 100);
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect);

//Fill shape with solid color
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.Yellow;

//Set the fill format of line
shape.Line.FillType = FillFormatType.Solid;
shape.Line.SolidFillColor.Color = Color.Gray;
```

---

# spire.presentation csharp find shape
## find shape by alternative text in presentation slide
```csharp
private IShape FindShape(ISlide slide, string altText)
{
    //Loop through shapes in the slide
    foreach (IShape shape in slide.Shapes)
    {
        //Find the shape whose alternative text is altText
        if (shape.AlternativeText.CompareTo(altText) == 0)
        {
            return shape;
        }
    }
    return null;
}
```

---

# Spire.Presentation C# Title Extraction
## Extract all title, centered title, and subtitle text from a PowerPoint presentation
```csharp
//Create a list to store title shapes
List<IShape> shapelist = new List<IShape>();

//Loop through all slides and all shapes on each slide
foreach (ISlide slide in ppt.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape.Placeholder != null)
        {
            //Get all titles based on placeholder type
            switch (shape.Placeholder.Type)
            {
                case PlaceholderType.Title:
                    shapelist.Add(shape);
                    break;
                case PlaceholderType.CenteredTitle:
                    shapelist.Add(shape);
                    break;
                case PlaceholderType.Subtitle:
                    shapelist.Add(shape);
                    break;
            }
        }
    }
}

//Extract text from all title shapes
StringBuilder sb = new StringBuilder();
sb.AppendLine("Below are all the obtained titles:");
for (int i = 0; i < shapelist.Count; i++)
{
    IAutoShape shape1 = shapelist[i] as IAutoShape;
    sb.AppendLine(shape1.TextFrame.Text);
}
```

---

# spire.presentation csharp animation
## get animation effect information from presentation
```csharp
// Create a PPT document
Presentation presentation = new Presentation();

// Load the document from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\..\Data\Animation.pptx");

// Travel each slide
foreach (ISlide slide in presentation.Slides)
{
    foreach (AnimationEffect effect in slide.Timeline.MainSequence)
    {
        // Get the animation effect type
        AnimationEffectType animationEffectType = effect.AnimationEffectType;
        
        // Get the slide number where the animation is located
        int slideNumber = slide.SlideNumber;
        
        // Get the shape name
        string shapeName = effect.ShapeTarget.Name;
    }
}
```

---

# spire.presentation animation motion path
## extract animation motion path data from powerpoint shapes
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
//Get the first shape
IShape shape = slide.Shapes[0];
//Create a StringBuilder to save the tracks
StringBuilder sb = new StringBuilder();
int i = 1;
//Traverse all animations
foreach (AnimationEffect effect in shape.Slide.Timeline.MainSequence)
{
    if (effect.ShapeTarget.Equals(shape as Shape))
    {
        //Get MotionPath
        MotionPath path = ((AnimationMotion)effect.CommonBehaviorCollection[0]).Path;
        //Get all points in the path
        foreach (MotionCmdPath motionCmdPath in path)
        {
            PointF[] points = motionCmdPath.Points;
            MotionCommandPathType type = motionCmdPath.CommandType;
            if (points != null)
            {
                foreach (PointF point in points)
                {
                    sb.AppendLine(i+"  MotionType: " + type + " -> X: " + point.X + ", Y: " + point.Y);
                }
                i++;
            }
        }
    }
}
```

---

# spire.presentation csharp text metrics
## get ascent and descent of text in presentation
```csharp
// Access the first slide in the presentation
ISlide slide = ppt.Slides[0];

// Access the first AutoShape in the slide
IAutoShape autoshape = slide.Shapes[0] as IAutoShape;

// Retrieve the layout lines from the TextFrame of the AutoShape
IList<LineText> lines = autoshape.TextFrame.GetLayoutLines();

// Iterate through each layout line
for (int i = 0; i < lines.Count; i++)
{
    // Get the ascent and descent properties of the current line
    float ascent = lines[i].Ascent;
    float descent = lines[i].Descent;
}
```

---

# Spire.Presentation C# Get Shape Group Alternative Text
## This code demonstrates how to extract alternative text from shapes within a group shape in a PowerPoint presentation.
```csharp
//Get alternative text from shape groups in a presentation
StringBuilder builder = new StringBuilder();

//Loop through slides and shapes
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is GroupShape)
        {
            //Find the shape group
            GroupShape groupShape = shape as GroupShape;
            foreach (IShape gShape in groupShape.Shapes)
            {
                //Append the alternative text in builder
                builder.AppendLine(gShape.AlternativeText);
            }
        }
    }
}
```

---

# spire.presentation csharp get shape points
## retrieve and display point information from a shape in a presentation
```csharp
//Get the first shape in first slide
IAutoShape shape = (IAutoShape)ppt.Slides[0].Shapes[0];

//Get the Point of shape
IList <PointF> points = shape.Points;

StringBuilder sb = new StringBuilder();
sb.Append("point count: " + points.Count + "\r\n");

for (int i = 0; i < points.Count; i++)
{
    sb.Append("point" + i + " " + points[i] + "\r\n");
}
```

---

# Spire Presentation Get Shapes by Placeholder
## Retrieve shapes from PowerPoint slides using placeholder and extract text
```csharp
//Get Placeholder
Placeholder placeholder = ppt.Slides[1].Shapes[0].Placeholder;
//Get Shapes by Placeholder
IShape[] shapes = ppt.Slides[1].GetPlaceholderShapes(placeholder);
string text = "";
//Iterate over all the shapes
for (int i = 0; i < shapes.Length; i++)
{
    //If shape is IAutoShape
    if (shapes[i] is IAutoShape)
    {
        IAutoShape autoShape = shapes[i] as IAutoShape;
        if (autoShape.TextFrame != null)
        {
            text += autoShape.TextFrame.Text + "\r\n";
        }
    }
}
```

---

# Spire.Presentation C# Get Text Frame Size
## Get the text frame size of shapes in a PowerPoint presentation
```csharp
// Get the first slide from presentation
ISlide slide = presentation.Slides[0];

// Iterate the shapes in the slide
for (int i = 0; i < slide.Shapes.Count; i++)
{
    IAutoShape autoShape = slide.Shapes[i] as IAutoShape;
    // Get the text frame size of the shape
    SizeF size = autoShape.TextFrame.GetTextSize();
    // The size contains width and height of the text frame
}
```

---

# Spire.Presentation C# Shape Text Extraction
## Extract text lines from shapes in a presentation slide
```csharp
// Get the first slide
ISlide slide = presentation.Slides[0];

// Iterate the shapes in the slide
for (int i=0;i<slide.Shapes.Count;i++)
{
    // Get shape 
    IAutoShape shape = (IAutoShape)slide.Shapes[i];
    
    // Get text lines in the shape and get the text
    IList<LineText> lines = shape.TextFrame.GetLayoutLines();
    for (int j = 0; j < lines.Count; j++)
    {
        // Process the text in each line
        string text = lines[j].Text;
    }
}
```

---

# spire.presentation csharp text position
## get text position within shapes in PowerPoint presentation
```csharp
// Access the first slide in the presentation
ISlide slide = ppt.Slides[0];

// Iterate through all the shapes in the slide
for (int i = 0; i < slide.Shapes.Count; i++)
{
    // Get the current shape
    IShape shape = slide.Shapes[i];

    // Check if the shape is an AutoShape
    if (shape is IAutoShape)
    {
        // Cast the shape to an AutoShape
        IAutoShape autoshape = slide.Shapes[i] as IAutoShape;

        // Get the text content of the AutoShape
        string text = autoshape.TextFrame.Text;

        // Obtain the text position information within the AutoShape
        PointF point = autoshape.TextFrame.GetTextLocation();

        // Append information about the shape, text, and location
        sb.AppendLine("Shape " + i + "：" + text + "\r\n" + "location：" + point.ToString());
    }
}
```

---

# Spire.Presentation C# Group Shapes
## Group shapes in a PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();
//Get the first slide
ISlide slide = ppt.Slides[0];

//Create two shapes in the slide
IShape rectangle = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 180, 200, 40));
rectangle.Fill.FillType = FillFormatType.Solid;
rectangle.Fill.SolidColor.KnownColor = KnownColors.SkyBlue;
rectangle.Line.Width = 0.1f;
IShape ribbon = slide.Shapes.AppendShape(ShapeType.Ribbon2, new RectangleF(290, 155, 120, 80));
ribbon.Fill.FillType = FillFormatType.Solid;
ribbon.Fill.SolidColor.KnownColor = KnownColors.LightPink;
ribbon.Line.Width = 0.1f;

//Add the two shape objects to an array list
ArrayList list = new ArrayList();
list.Add(rectangle);
list.Add(ribbon);

//Group the shapes in the list
ppt.Slides[0].GroupShapes(list);
```

---

# Spire.Presentation C# Hide Shape
## Hide a specific shape in a PowerPoint presentation by its alternative text
```csharp
//Loop through slides
foreach (ISlide slide in presentation.Slides)
{
    //Loop through shapes in the slide
    foreach (IShape shape in slide.Shapes)
    {
        //Find the shape whose alternative text is Shape1
        if (shape.AlternativeText.CompareTo("Shape1") == 0)
        {
            //Hide the shape
            shape.IsHidden = true;
        }
    }
}
```

---

# Spire.Presentation C# TextBox Detection
## Determine if a PowerPoint shape is a textbox
```csharp
// Iterate through all slides in the presentation
foreach (ISlide slide in presentation.Slides)
{
    // Iterate through all shapes in each slide
    foreach (IShape shape in slide.Shapes)
    {
        // Check if the shape is an IAutoShape
        if (shape is IAutoShape)
        {
            // Determine if the shape is a textbox
            Boolean isTextbox = shape.IsTextBox;
            // Store the result
            string result = isTextbox ? "shape is text box" : "shape is not text box";
        }
    }
}
```

---

# spire presentation placeholder operations
## operate on different types of placeholders in presentation slides
```csharp
// Operate placeholders
for (int j = 0; j < presentation.Slides.Count; j++)
{
    ISlide slide = (ISlide)presentation.Slides[j];
    
    for (int i = 0; i < slide.Shapes.Count; i++)
    {
        Shape shape = (Shape)slide.Shapes[i];
        switch (shape.Placeholder.Type)
        {
            case PlaceholderType.Media:
                shape.InsertVideo("Video.mp4");
                break;
           
            case PlaceholderType.Picture:
                shape.InsertPicture("E-iceblueLogo.png");
                break;
            
            case PlaceholderType.Chart:
                shape.InsertChart(ChartType.ColumnClustered);
                break;
            
            case PlaceholderType.Table:
                shape.InsertTable(3, 2);
                break;
            
            case PlaceholderType.Diagram:
                shape.InsertSmartArt(SmartArtLayoutType.BasicBlockList);
                break;
        }
    }
}
```

---

# Spire.Presentation C# Shape Locking
## Prevent or allow changing shapes in PowerPoint presentations
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Add a rectangle shape to the slide
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, 400, 150));

//The changes of selection and rotation are allowed
shape.Locking.RotationProtection = false;
shape.Locking.SelectionProtection = false;
//The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed 
shape.Locking.ResizeProtection = true;
shape.Locking.PositionProtection = true;
shape.Locking.ShapeTypeProtection = true;
shape.Locking.AspectRatioProtection = true;
shape.Locking.TextEditingProtection = true;
shape.Locking.AdjustHandlesProtection = true;
```

---

# spire.presentation csharp remove shapes
## Remove shapes from PowerPoint presentation based on alternative text
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Loop through slides
for (int i = 0; i < presentation.Slides.Count; i++)
{
    ISlide slide = presentation.Slides[i];
    //Loop through shapes
    for (int j = 0; j < slide.Shapes.Count; j++)
    {
        IShape shape = slide.Shapes[j];
        //Find the shapes whose alternative text contain "Shape"
        if(shape.AlternativeText.Contains("Shape"))
        {
            slide.Shapes.Remove(shape);
            j--;
        }
    }
}
```

---

# Spire.Presentation C# Shape Reordering
## Change the Z-order (stacking order) of overlapping shapes in a PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first shape of the first slide
IShape shape = ppt.Slides[0].Shapes[0];
//Change the shape's zorder
ppt.Slides[0].Shapes.ZOrder(1, shape);
```

---

# spire.presentation csharp placeholder
## reset position of placeholders in PowerPoint presentation
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_7.pptx");

//Get the first slide from the sample document.
ISlide slide = presentation.Slides[0];

foreach (IShape shapeToMove in slide.Shapes)
{
    //Reset the position of the slide number to the left.
    if (shapeToMove.Name.Contains("Slide Number Placeholder"))
    {
        shapeToMove.Left = 0;
    }

    else if (shapeToMove.Name.Contains("Date Placeholder"))
    {
        //Reset the position of the date time to the center.
        shapeToMove.Left = presentation.SlideSize.Size.Width / 2;

        //Reset the date time display style.
        (shapeToMove as IAutoShape).TextFrame.TextRange.Paragraph.Text = DateTime.Now.ToString("dd.MM.yyyy");
        (shapeToMove as IAutoShape).TextFrame.IsCentered = true;
    }
}
```

---

# Spire.Presentation Shape Resizing
## Resize and reposition shapes when changing slide size
```csharp
//Define the original slide size
float currentHeight = ppt.SlideSize.Size.Height;
float currentWidth = ppt.SlideSize.Size.Width;

//Change the slide size as A3
ppt.SlideSize.Type = SlideSizeType.A3;

//Define the new slide size
float newHeight = ppt.SlideSize.Size.Height;
float newWidth = ppt.SlideSize.Size.Width;

//Define the ratio from the old and new slide size
float ratioHeight = newHeight / currentHeight;
float ratioWidth = newWidth / currentWidth;

//Reset the size and position of the shape on the slide
foreach (ISlide slide in ppt.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        shape.Height = shape.Height * ratioHeight;
        shape.Width = shape.Width * ratioWidth;

        shape.Left = shape.Left * ratioHeight;
        shape.Top = shape.Top * ratioWidth;
    }
}
```

---

# spire.presentation csharp rotate shapes
## rotate shapes in powerpoint presentation
```csharp
//Get the shapes 
IAutoShape shape = ppt.Slides[0].Shapes[0] as IAutoShape;

//Set the rotation
shape.Rotation = 60;

(ppt.Slides[0].Shapes[1] as IAutoShape).Rotation = 120;
(ppt.Slides[0].Shapes[2] as IAutoShape).Rotation = 180;
(ppt.Slides[0].Shapes[3] as IAutoShape).Rotation = 240;
```

---

# spire.presentation csharp save shape as svg
## save presentation shapes as svg files
```csharp
// Get the first slide
ISlide slide = presentation.Slides[0];

// Iterate the shapes in the slide
for (int i = 0; i < slide.Shapes.Count; i++)
{
    // Save the shapes as SVG
    byte[] svgByte = slide.Shapes[i].SaveAsSvgInSlide();
    // Create file stream and write SVG data
    FileStream fs = new FileStream("shapePath_" + i + ".svg", FileMode.Create);
    fs.Write(svgByte, 0, svgByte.Length);
    fs.Close();
}
```

---

# spire.presentation csharp 3d effects
## set 3D effects for shapes in presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Add shape1 and fill it with color
IAutoShape shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(150, 150, 150, 150));
shape1.Fill.FillType = FillFormatType.Solid;
shape1.Fill.SolidColor.KnownColor = KnownColors.SkyBlue;
//Initialize a new instance of the 3-D class for shape1 and set its properties
ShapeThreeD effect1 = shape1.ThreeD.ShapeThreeD;
effect1.PresetMaterial = PresetMaterialType.Powder;
effect1.TopBevel.PresetType = BevelPresetType.ArtDeco;
effect1.TopBevel.Height = 4;
effect1.TopBevel.Width = 12;
effect1.BevelColorMode = BevelColorType.Contour;
effect1.ContourColor.KnownColor = KnownColors.LightBlue;
effect1.ContourWidth = 3.5;

//Add shape2 and fill it with color
IAutoShape shape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Pentagon, new RectangleF(400, 150, 150, 150));
shape2.Fill.FillType = FillFormatType.Solid;
shape2.Fill.SolidColor.KnownColor = KnownColors.LightGreen;
//Initialize a new instance of the 3-D class for shape2 and set its properties
ShapeThreeD effect2 = shape2.ThreeD.ShapeThreeD;
effect2.PresetMaterial = PresetMaterialType.SoftEdge;
effect2.TopBevel.PresetType = BevelPresetType.SoftRound;
effect2.TopBevel.Height = 12;
effect2.TopBevel.Width = 12;
effect2.BevelColorMode = BevelColorType.Contour;
effect2.ContourColor.KnownColor = KnownColors.LawnGreen;
effect2.ContourWidth = 5;
```

---

# Spire.Presentation C# Alternative Text
## Set and get alternative text for shapes in PowerPoint presentations
```csharp
//Set the alternative text (title and description)
slide.Shapes[0].AlternativeTitle = "Rectangle";
slide.Shapes[0].AlternativeText = "This is a Rectangle";

//Get the alternative text (title and description)
string alternativeText = null;
string title = slide.Shapes[0].AlternativeTitle;
alternativeText += "Title: " + title + "\r\n";
string description = slide.Shapes[0].AlternativeText;
alternativeText += "Description: " + description;
```

---

# spire.presentation csharp animation
## set animation type and time value for text animation
```csharp
//Set the AnimateType as Letter
ppt.Slides[0].Timeline.MainSequence[0].IterateType = Spire.Presentation.Drawing.TimeLine.AnimateType.Letter;

//Set the IterateTimeValue for the animate text
ppt.Slides[0].Timeline.MainSequence[0].IterateTimeValue = 10;
```

---

# spire.presentation csharp animation
## set animation repeat type for PowerPoint presentations
```csharp
//Create PPT document
Presentation presentation = new Presentation();
//Get the first slide
ISlide slide = presentation.Slides[0];
AnimationEffectCollection animations = slide.Timeline.MainSequence;

animations[0].Timing.AnimationRepeatType = AnimationRepeatType.UtilEndOfSlide;
```

---

# spire.presentation csharp gradient stop
## set brightness and transparency for gradient stop
```csharp
// Set the color of shape
shape.Fill.FillType = FillFormatType.Gradient;

// Add gradient stops to create a gradient fill
shape.Fill.Gradient.GradientStops.Append(0f, KnownColors.Olive);
shape.Fill.Gradient.GradientStops.Append(1f, KnownColors.PowderBlue);

// Adjust the brightness and transparency of the first gradient stop
shape.Fill.Gradient.GradientStops[0].Color.Brightness = 0.5f;
shape.Fill.Gradient.GradientStops[0].Color.Transparency = 0.5f;
```

---

# spire.presentation csharp ellipse
## Set ellipse format in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];

//Add a rectangle
RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 100, 100, 200, 100);
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Ellipse, rect);

//Set the fill format of shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.CadetBlue;

//Set the fill format of line
shape.Line.FillType = FillFormatType.Solid;
shape.Line.SolidFillColor.Color = Color.DimGray;
```

---

# spire.presentation csharp line formatting
## Set format for lines in presentation shapes
```csharp
//Add a rectangle shape to the slide
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 150, 200, 100));
//Set the fill color of the rectangle shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
//Apply some formatting on the line of the rectangle
shape.Line.Style = TextLineStyle.ThickThin;
shape.Line.Width = 5;
shape.Line.DashStyle = LineDashStyleType.Dash;
//Set the color of the line of the rectangle
shape.ShapeStyle.LineColor.Color = Color.SkyBlue;

//Add a ellipse shape to the slide
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(400, 150, 200, 100));
//Set the fill color of the ellipse shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.White;
//Apply some formatting on the line of the ellipse
shape.Line.Style = TextLineStyle.ThickBetweenThin;
shape.Line.Width = 5;
shape.Line.DashStyle = LineDashStyleType.DashDot;
//Set the color of the line of the ellipse
shape.ShapeStyle.LineColor.Color = Color.OrangeRed;
```

---

# spire.presentation csharp shapes
## set line join styles for shapes
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.Slides[0];

//Add three shapes
IAutoShape shape1 = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 150, 150, 50));
IAutoShape shape2 = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(250, 150, 150, 50));
IAutoShape shape3 = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(450, 150, 150, 50));

//Fill shapes
shape1.Fill.FillType = FillFormatType.Solid;
shape1.Fill.SolidColor.Color = Color.CadetBlue;
shape2.Fill.FillType = FillFormatType.Solid;
shape2.Fill.SolidColor.Color = Color.CadetBlue;
shape3.Fill.FillType = FillFormatType.Solid;
shape3.Fill.SolidColor.Color = Color.CadetBlue;

//Fill lines of shapes
shape1.Line.FillType = FillFormatType.Solid;
shape1.Line.SolidFillColor.Color = Color.DarkGray;
shape2.Line.FillType = FillFormatType.Solid;
shape2.Line.SolidFillColor.Color = Color.DarkGray;
shape3.Line.FillType = FillFormatType.Solid;
shape3.Line.SolidFillColor.Color = Color.DarkGray;

//Set the line width
shape1.Line.Width = 10;
shape2.Line.Width = 10;
shape3.Line.Width = 10;

//Set the join styles of lines
shape1.Line.JoinStyle = LineJoinType.Bevel;
shape2.Line.JoinStyle = LineJoinType.Miter;
shape3.Line.JoinStyle = LineJoinType.Round;

//Add text in shapes
shape1.TextFrame.Text = "Bevel Join Style";
shape2.TextFrame.Text = "Miter Join Style";
shape3.TextFrame.Text = "Round Join Style";
```

---

# spire.presentation csharp shape effects
## set outline and effect for shapes in presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Draw a Rectangle shape
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(150, 180, 100, 50));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.SkyBlue;
//Set outline color
shape.ShapeStyle.LineColor.Color = Color.Red;
//Set shadow effect
PresetShadow shadow = new PresetShadow();
shadow.ColorFormat.Color = Color.LightSkyBlue;
shadow.Preset = PresetShadowValue.FrontRightPerspective;
shadow.Distance = 10.0;
shadow.Direction = 225.0f;
shape.EffectDag.PresetShadowEffect = shadow;

//Draw a Ellipse shape
shape = slide.Shapes.AppendShape(ShapeType.Ellipse, new RectangleF(400, 150, 100, 100));
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.SkyBlue;
//Set outline color
shape.ShapeStyle.LineColor.Color = Color.Yellow;
//Set glow effect
GlowEffect glow = new GlowEffect();
glow.ColorFormat.Color = Color.LightPink;
glow.Radius = 20.0;
shape.EffectDag.GlowEffect = glow;
```

---

# spire.presentation csharp rounded rectangle
## set radius for different types of rounded rectangles in a PowerPoint slide
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide 
ISlide slide = presentation.Slides[0];

//Insert a rectangle with four round corners and set its radius
IAutoShape shape1 = slide.Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 50, 150, 150));
shape1.SetRoundRadius(shape1.Width / 3);

//Insert a rectangle with one round corner and set its radius
IAutoShape shape2 = slide.Shapes.AppendShape(ShapeType.OneRoundCornerRectangle, new RectangleF(250, 50, 150, 150));
shape2.SetRoundRadius(shape2.Width / 3);

//Insert a rectangle with one round corner and which one round cornet is snipped and set its radius
IAutoShape shape3 = slide.Shapes.AppendShape(ShapeType.OneSnipOneRoundCornerRectangle, new RectangleF(450, 50, 150, 150));
shape3.SetRoundRadius(shape3.Width / 3);

//Insert a rectangle with two diagonal round corners and set its radius
IAutoShape shape4 = slide.Shapes.AppendShape(ShapeType.TwoDiagonalRoundCornerRectangle, new RectangleF(50, 250, 150, 150));
shape4.SetRoundRadius(shape4.Width / 3);

//Insert a rectangle with two same side round corners and set its radius
IAutoShape shape5 = slide.Shapes.AppendShape(ShapeType.TwoSamesideRoundCornerRectangle, new RectangleF(250, 250, 150, 150));
shape5.SetRoundRadius(shape5.Width / 3);
```

---

# spire.presentation csharp rounded rectangle
## set radius of rounded rectangle in presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Insert a rounded rectangle and set its radious
presentation.Slides[0].Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10);

//Append a rounded rectangle and set its radius
IAutoShape shape = presentation.Slides[0].Shapes.AppendRoundRectangle(380, 180, 100, 200, 100);
//Set the color and fill style of shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.SeaGreen;
shape.ShapeStyle.LineColor.Color = Color.White;

//Rotate the shape to 90 degree
shape.Rotation = 90;
```

---

# spire.presentation csharp shape formatting
## set rectangle fill and line format
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add a shape
RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 100, 100, 200, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect);

//Set the fill format of shape
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.CadetBlue;

//Set the fill format of line
shape.Line.FillType = FillFormatType.Solid;
shape.Line.SolidFillColor.Color = Color.DimGray;
```

---

# spire.presentation csharp shadow effect
## Set inner shadow effect for shape in presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

ISlide slide = ppt.Slides[0];

//Add a shape to slide.
RectangleF rect1 = new RectangleF(200, 150, 300, 120);
IAutoShape shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect1);
shape.Fill.FillType = FillFormatType.Solid;
shape.Fill.SolidColor.Color = Color.LightBlue;
shape.Line.FillType = FillFormatType.None;
shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape.";
shape.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid;
shape.TextFrame.TextRange.Fill.SolidColor.Color = Color.Black;

//Create an inner shadow effect through InnerShadowEffect object. 
InnerShadowEffect innerShadow = new InnerShadowEffect();
innerShadow.BlurRadius = 20;
innerShadow.Direction = 0;
innerShadow.Distance = 0;
innerShadow.ColorFormat.Color = Color.Black;

//Apply the shadow effect to shape.
shape.EffectDag.InnerShadowEffect = innerShadow;
```

---

# spire.presentation csharp shape to image
## Convert presentation shapes to image files
```csharp
//Save shapes as images
Image image = presentation.Slides[0].Shapes[i].SaveAsImage();

//The following method also can save shape as image
//Image image = presentation.Slides[0].Shapes.SaveAsImage(i);

//Write image to Png
string fileName = String.Format("Picture-{0}.png", i);
image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
```

---

# spire.presentation csharp shape conversion
## convert presentation shapes to SVG format
```csharp
// Create a new Presentation object
Presentation ppt = new Presentation();

// Load a PowerPoint file
ppt.LoadFromFile("presentation.pptx");

// Access the first slide in the presentation
ISlide slide = ppt.Slides[0];

// Iterate through each shape in the slide
foreach (IShape shape in slide.Shapes)
{
    // Save the shape as SVG format
    byte[] svgByte = shape.SaveAsSvg();
    
    // Further processing with SVG data would go here
}

// Dispose of the Presentation object to release resources
ppt.Dispose();
```

---

# spire.presentation csharp ungroup shapes
## ungroup shapes in presentation slides
```csharp
//Get the GroupShape
GroupShape groupShape = ppt.Slides[0].Shapes[0] as GroupShape;
//Ungroup the shapes
ppt.Slides[0].Ungroup(groupShape);
```

---

# spire.presentation csharp section
## add sections to powerpoint presentation
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//Get the second slide
ISlide slide = ppt.Slides[1];

//Append section with section name at the end
ppt.SectionList.Append("E-iceblue01");
//Add section with slide
ppt.SectionList.Add("section1", slide);
```

---

# Spire.Presentation C# Section Management
## Add slide to a new section in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add a new shape to the PPT document
presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(200, 50, 300, 100));

//Create a new section and copy the first slide to it
Section NewSection = presentation.SectionList.Append("New Section");
NewSection.Insert(0, presentation.Slides[0]);
```

---

# Spire.Presentation C# Section Management
## Delete sections from a PowerPoint presentation
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//remove the specified section
//ppt.SectionList.RemoveAt(3);
//remove all sections
ppt.SectionList.RemoveAll();
```

---

# spire.presentation csharp section index
## Get section index from PowerPoint presentation
```csharp
Section section = ppt.SectionList[0];
int index = ppt.SectionList.IndexOf(section);
```

---

# spire.presentation encrypted stream
## load encrypted PowerPoint presentation from stream
```csharp
// Create a Presentation instance
Presentation ppt = new Presentation();

//Load PowerPoint file from stream
FileStream from_stream = File.OpenRead(@"..\..\..\..\..\..\Data\\OpenEncryptedPPT.pptx");

// The password
String password = "123456";

// Load the encrypted stream with the provided password
ppt.LoadFromStream(from_stream, FileFormat.Auto, password);
```

---

# spire.presentation csharp load from stream
## load PowerPoint presentation from stream
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Load PowerPoint file from stream
FileStream from_stream = File.OpenRead(@"..\..\..\..\..\..\Data\InputTemplate.pptx");
ppt.LoadFromStream(from_stream, FileFormat.Pptx2013);        

//Save the document
string result = "LoadFromStream.pptx";
ppt.SaveToFile(result, FileFormat.Pptx2013);
from_stream.Dispose();
```

---

# spire.presentation csharp loop presentation
## configure PowerPoint presentation to loop continuously
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Set the Boolean value of ShowLoop as true
ppt.ShowLoop = true;

//Set the PowerPoint document to show animation and narration
ppt.ShowAnimation = true;
ppt.ShowNarration = true;
//Use slide transition timings to advance slide
ppt.UseTimings = true;
```

---

# Spire.Presentation C# Page Setup
## Configure slide size, orientation and type in PowerPoint presentation
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Set the size of slides
presentation.SlideSize.Size = new SizeF(600,600);
presentation.SlideSize.Orientation = SlideOrienation.Portrait;
presentation.SlideSize.Type = SlideSizeType.Custom;
```

---

# Spire.Presentation C# Save to Stream
## Demonstrates how to create a PowerPoint presentation and save it to a stream
```csharp
//Create PowerPoint file
Presentation presentation = new Presentation();

//Append new shape
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 100, 600, 150));
shape.Fill.FillType = FillFormatType.None;
shape.ShapeStyle.LineColor.Color = Color.White;

//Add text to shape
shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream.";
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid;
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 30;

//Save to Stream
FileStream to_stream = new FileStream("SaveToStream.pptx", FileMode.Create);
presentation.SaveToFile(to_stream, FileFormat.Pptx2013);
to_stream.Close();
```

---

# spire.presentation csharp kiosk mode
## set presentation show type as kiosk
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Specify the presentation show type as kiosk
ppt.ShowType = SlideShowType.Kiosk;
```

---

# Spire.Presentation C# Split PowerPoint
## Split a PowerPoint presentation into individual slides
```csharp
//Create a new presentation and load the original file
Presentation ppt = new Presentation();
ppt.LoadFromFile("InputTemplate.pptx");

//Loop through each slide
for (int i = 0; i < ppt.Slides.Count; i++)
{
    //Create a new presentation, remove the default blank slide
    Presentation newppt = new Presentation();
    newppt.Slides.RemoveAt(0);

    //Append the current slide to the new presentation
    newppt.Slides.Append(ppt.Slides[i]);

    //Save the new presentation with just that one slide
    string result = string.Format("SplitPPT-{0}.pptx", i);
    newppt.SaveToFile(result, FileFormat.Pptx2010);
}
```

---

# Spire.Presentation C# Get Built-in Properties
## Retrieve built-in document properties from a PowerPoint presentation
```csharp
// Create PPT document
Presentation presentation = new Presentation();

// Load the PPT document from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\GetProperties.pptx");

// Get the builtin properties 
string application = presentation.DocumentProperty.Application;
string author = presentation.DocumentProperty.Author;
string company = presentation.DocumentProperty.Company;
string keywords = presentation.DocumentProperty.Keywords;
string comments = presentation.DocumentProperty.Comments;
string category = presentation.DocumentProperty.Category;
string title = presentation.DocumentProperty.Title;
string subject = presentation.DocumentProperty.Subject;
```

---

# Spire.Presentation C# Mark As Final
## Mark a PowerPoint presentation as final to prevent further editing
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile("MarkAsFinal.pptx");

//Mark the document as final
presentation.DocumentProperty["_MarkAsFinal"] = true;

//Save the document
presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);
```

---

# spire.presentation csharp document properties
## set presentation document properties
```csharp
//Set the DocumentProperty of PPT document
presentation.DocumentProperty.Application = "Spire.Presentation";
presentation.DocumentProperty.Author = "E-iceblue";
presentation.DocumentProperty.Company = "E-iceblue Co., Ltd.";
presentation.DocumentProperty.Keywords = "Demo File";
presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
presentation.DocumentProperty.Category = "Demo";
presentation.DocumentProperty.Title = "This is a demo file.";
presentation.DocumentProperty.Subject = "Test";
```

---

# Spire.Presentation Set Template Properties
## Set document properties for presentation templates
```csharp
// Create a presentation
Presentation presentation = new Presentation();

// Set the DocumentProperty 
presentation.DocumentProperty.Application = "Spire.Presentation";
presentation.DocumentProperty.Author = "E-iceblue";
presentation.DocumentProperty.Company = "E-iceblue Co., Ltd.";
presentation.DocumentProperty.Keywords = "Demo File";
presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation.";
presentation.DocumentProperty.Category = "Demo";
presentation.DocumentProperty.Title = "This is a demo file.";
presentation.DocumentProperty.Subject = "Test";

// Save to template file
presentation.SaveToFile(filePath, fileFormat);
```

---

# Spire.Presentation Digital Signature
## Add digital signature to PowerPoint presentation
```csharp
//Load a ppt document
Presentation ppt = new Presentation();
ppt.LoadFromFile("AddDigitalSignature.pptx");

//Load the certificate
X509Certificate2 x509 = new X509Certificate2("gary.pfx", "e-iceblue");

//Add a digital signature
ppt.AddDigitalSignature(x509, "111", DateTime.Now);

//Save the document
ppt.SaveToFile("AddDigitalSignature_result.pptx", FileFormat.Pptx2010);
```

---

# Spire.Presentation C# Password Protection Check
## Check if a PowerPoint presentation is password protected
```csharp
// Create Presentation
Presentation presentation = new Presentation();

// Check whether a PPT document is password protected
bool isProtected = presentation.IsPasswordProtected("file_path.pptx");
```

---

# Spire.Presentation C# Encryption
## Encrypt a PowerPoint presentation with password protection
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Encrypt.pptx");

//Get the password that the user entered
string password = this.textBox1.Text;

//Encrypy the document with the password
presentation.Encrypt(password);

//Save the document
presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);
```

---

# Spire.Presentation C# Password Modification
## Modify password of an encrypted PowerPoint presentation
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the encrypted file with original password.
presentation.LoadFromFile(filePath, "123456");

//Remove the encryption.
presentation.RemoveEncryption();

//Protect the document by setting a new password.
presentation.Protect("654321");

//Save the modified file.
presentation.SaveToFile(result, FileFormat.Pptx2013);
```

---

# spire.presentation csharp encrypted ppt
## open encrypted PowerPoint presentation with password
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the PPT with password
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\OpenEncryptedPPT.pptx", FileFormat.Pptx2010, textBox1.Text);

//Save as a new PPT with original password
presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);

//Launch the PPT file
System.Diagnostics.Process.Start("Output.pptx");
```

---

# spire.presentation csharp security
## remove all digital signatures from presentation
```csharp
//Remove all digital signatures
if (ppt.IsDigitallySigned == true)
{
    ppt.RemoveAllDigitalSignatures();
}
```

---

# spire.presentation csharp encryption
## remove encryption from PowerPoint presentation
```csharp
//Create a PowerPoint document
Presentation presentation = new Presentation();

//Load the encrypted file from disk with password
presentation.LoadFromFile("file_path", "password");

//Remove encryption
presentation.RemoveEncryption();

//Save to file
presentation.SaveToFile("result_path", FileFormat.Pptx2013);
```

---

# spire.presentation csharp security
## set presentation document to read-only with password protection
```csharp
//Load a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile("SetDocumentReadOnly.pptx");

//Get the password that the user entered
string password = textBox1.Text;

//Protect the document with the password
presentation.Protect(password);

//Save the document
presentation.SaveToFile("Output.pptx", FileFormat.Pptx2010);
```

---

# spire.presentation csharp background
## set slide background with different types
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Set the background of the first slide to Gradient color
presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom;
presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Gradient;
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear;
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStyle = Spire.Presentation.Drawing.GradientStyle.FromCorner1;
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.Append(1f, KnownColors.SkyBlue);
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.Append(0f, KnownColors.White);

//Set the background of the second slide to Solid color
presentation.Slides[1].SlideBackground.Type = BackgroundType.Custom;
presentation.Slides[1].SlideBackground.Fill.FillType = FillFormatType.Solid;
presentation.Slides[1].SlideBackground.Fill.SolidColor.Color = Color.SkyBlue;

presentation.Slides.Append();
//Set the background of the third slide to picture
RectangleF rect = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
presentation.Slides[2].SlideBackground.Fill.FillType = FillFormatType.Picture;
IEmbedImage image = presentation.Slides[2].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
presentation.Slides[2].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image as IImageData;
```

---

# spire.presentation csharp gradient background
## Set gradient background for presentation slide
```csharp
//Set the background to gradient
slide.SlideBackground.Type = BackgroundType.Custom;
slide.SlideBackground.Fill.FillType = FillFormatType.Gradient;

//Add gradient stops
slide.SlideBackground.Fill.Gradient.GradientStops.Append(0.1f, Color.LightSeaGreen);
slide.SlideBackground.Fill.Gradient.GradientStops.Append(0.7f, Color.LightCyan);

//Set gradient shape type
slide.SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear;

//Set the angle
slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45;
```

---

# Spire.Presentation Master Background
## Setting master slide background in PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Set the slide background of master
presentation.Masters[0].SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom;
presentation.Masters[0].SlideBackground.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
presentation.Masters[0].SlideBackground.Fill.SolidColor.Color = Color.LightSalmon;
```

---

# spire.presentation csharp error bars
## add and format error bars in powerpoint charts
```csharp
//Get the column chart on the first slide and set chart title.
IChart columnChart = presentation.Slides[0].Shapes[0] as IChart;
columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars"; 

//Add Y (Vertical) Error Bars.
//Get Y error bars of the first chart series.
IErrorBarsFormat errorBarsYFormat1 = columnChart.Series[0].ErrorBarsYFormat;

//Set end cap.
errorBarsYFormat1.ErrorBarNoEndCap = false;

//Specify direction.
errorBarsYFormat1.ErrorBarSimType = ErrorBarSimpleType.Plus;

//Specify error amount type.
errorBarsYFormat1.ErrorBarvType = ErrorValueType.StandardError;

//Set value.
errorBarsYFormat1.ErrorBarVal = 0.3f;

//Set line format.
errorBarsYFormat1.Line.FillType = FillFormatType.Solid;
errorBarsYFormat1.Line.SolidFillColor.Color = Color.MediumVioletRed;
errorBarsYFormat1.Line.Width = 1;

//Get the bubble chart on the second slide and set chart title.
IChart bubbleChart = presentation.Slides[1].Shapes[0] as IChart;
bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars";

//Add X (Horizontal) and Y (Vertical) Error Bars.
//Get X error bars of the first chart series.
IErrorBarsFormat errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat;

//Set end cap.
errorBarsXFormat.ErrorBarNoEndCap = false;

//Specify direction.
errorBarsXFormat.ErrorBarSimType = ErrorBarSimpleType.Both;

//Specify error amount type.
errorBarsXFormat.ErrorBarvType = ErrorValueType.StandardError;

//Set value.
errorBarsXFormat.ErrorBarVal = 0.3f;

//Get Y error bars of the first chart series.
IErrorBarsFormat errorBarsYFormat2 = bubbleChart.Series[0].ErrorBarsYFormat;

//Set end cap.
errorBarsYFormat2.ErrorBarNoEndCap = false;

//Specify direction.
errorBarsYFormat2.ErrorBarSimType = ErrorBarSimpleType.Both;

//Specify error amount type.
errorBarsYFormat2.ErrorBarvType = ErrorValueType.StandardError;

//Set value.
errorBarsYFormat2.ErrorBarVal = 0.3f;
```

---

# Spire.Presentation C# Chart Error Bars
## Add custom error bars to chart series in PowerPoint presentation
```csharp
//Get X error bars of the first chart series
IErrorBarsFormat errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat;
//Specify error amount type as custom error bars
errorBarsXFormat.ErrorBarvType = ErrorValueType.CustomErrorBars;
//Set the minus and plus value of the X error bars
errorBarsXFormat.MinusVal = 0.5f;
errorBarsXFormat.PlusVal = 0.5f;

//Get Y error bars of the first chart series
IErrorBarsFormat errorBarsYFormat = bubbleChart.Series[0].ErrorBarsYFormat;
//Specify error amount type as custom error bars
errorBarsYFormat.ErrorBarvType = ErrorValueType.CustomErrorBars;
//Set the minus and plus value of the Y error bars
errorBarsYFormat.MinusVal = 1f;
errorBarsYFormat.PlusVal = 1f;
```

---

# spire.presentation csharp chart
## add secondary value axis to chart
```csharp
//Get the chart from the PowerPoint file.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Add a secondary axis to display the value of Series 3.
chart.Series[2].UseSecondAxis = true;

//Set the grid line of secondary axis as invisible.
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;
```

---

# spire.presentation csharp chart
## add shadow effect to data labels
```csharp
//Get the chart.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Add a data label to the first chart series.
ChartDataLabelCollection dataLabels = chart.Series[0].DataLabels;
ChartDataLabel Label = dataLabels.Add();
Label.LabelValueVisible = true;

//Add outer shadow effect to the data label.
Label.Effect.OuterShadowEffect = new OuterShadowEffect();

//Set shadow color.
Label.Effect.OuterShadowEffect.ColorFormat.Color = Color.Yellow;

//Set blur.
Label.Effect.OuterShadowEffect.BlurRadius = 5;

//Set distance.
Label.Effect.OuterShadowEffect.Distance = 10;

//Set angle.
Label.Effect.OuterShadowEffect.Direction = 90f;
```

---

# spire.presentation csharp chart
## add trendline for chart series
```csharp
//Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;
ITrendlines it = chart.Series[0].AddTrendLine(TrendlinesType.Linear);

//Set the trendline properties to determine what should be displayed.
it.displayEquation = false;
it.displayRSquaredValue = false;
```

---

# spire.presentation csharp pie chart
## create pie chart with auto vary color setting
```csharp
//Create a PPT file
Presentation ppt = new Presentation();

RectangleF rect1 = new RectangleF(40, 100, 550, 320);

//Add a pie chart
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Pie, rect1, false);
chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Attach the data to chart
string[] quarters = new string[] { "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr" };
int[] sales = new int[] { 210, 320, 180, 500 };
chart.ChartData[0, 0].Text = "Quarters";
chart.ChartData[0, 1].Text = "Sales";
for (int i = 0; i < quarters.Length; ++i)
{
    chart.ChartData[i + 1, 0].Value = quarters[i];
    chart.ChartData[i + 1, 1].Value = sales[i];
}

chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];

//Set whether auto vary color, default value is true
chart.Series[0].IsVaryColor = false;

chart.Series[0].Distance = 15;
```

---

# Spire.Presentation C# Chart Legend
## Change the color and style of a chart legend in a PowerPoint presentation
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Change the fill color
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.Color = Color.Blue;
//Use italic for the paragraph
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.IsItalic = TriState.True;
```

---

# spire.presentation csharp chart datatable
## change font for chart data table
```csharp
//Enable data table for chart
chart.HasDataTable = true;

//Add a new paragraph in data table
chart.ChartDataTable.Text.Paragraphs.Append(new TextParagraph());
//Change the font size
chart.ChartDataTable.Text.Paragraphs[0].DefaultCharacterProperties.FontHeight = 15;
```

---

# spire.presentation csharp chart legend
## change font size for chart legend
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Change legend font size
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 17;
```

---

# spire.presentation csharp chart
## change chart series name
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Get the ranges of series label 
CellRanges cr = Chart.Series.SeriesLabel;

//Change the value
cr[0].Value = "Changed series name";
```

---

# spire.presentation csharp trendline
## modify trendline equation properties in PowerPoint chart
```csharp
//Get chart on the first slide
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Get the first trendline 
ITrendlines trendline = chart.Series[0].TrendLines[0] as ITrendlines;

//Change font size for trendline Equation text
foreach (TextParagraph para in trendline.TrendLineLabel.TextFrameProperties.Paragraphs)
{
    para.DefaultCharacterProperties.FontHeight = 20;
    foreach (Spire.Presentation.TextRange range in para.TextRanges)
    {
        range.FontHeight = 20;
    }
}

//Change position for trendline Equation
trendline.TrendLineLabel.OffsetX = -0.1f;
trendline.TrendLineLabel.OffsetY = -0.05f;
```

---

# spire.presentation csharp chart
## change text font in chart
```csharp
//Get the chart
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Change the font of title
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Lucida Sans Unicode");
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue;
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 30;

//Change the font of legend
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.DarkGreen;
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Lucida Sans Unicode");

//Change the font of series
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Lucida Sans Unicode");
```

---

# spire.presentation csharp chart axis
## configure chart primary and secondary axis settings
```csharp
//Get the chart
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Add a secondary axis to display the value of Series 3
chart.Series[2].UseSecondAxis = true;

//Set the grid line of secondary axis as invisible
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;

//Set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
chart.PrimaryValueAxis.IsAutoMax = false;
chart.PrimaryValueAxis.IsAutoMin = false;
chart.SecondaryValueAxis.IsAutoMax = false;
chart.SecondaryValueAxis.IsAutoMax = false;

chart.PrimaryValueAxis.MinValue = 0f;
chart.PrimaryValueAxis.MaxValue = 5.0f;
chart.SecondaryValueAxis.MinValue = 0f;
chart.SecondaryValueAxis.MaxValue = 1.0f;

//Set axis line format
chart.PrimaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid;
chart.SecondaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid;
chart.PrimaryValueAxis.MinorGridLines.Width = 0.1f;
chart.SecondaryValueAxis.MinorGridLines.Width = 0.1f;
chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.LightGray;
chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.LightGray;
chart.PrimaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash;
chart.SecondaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash;

chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3f;
chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.LightSkyBlue;
chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3f;
chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.LightSkyBlue;
```

---

# spire.presentation csharp chart
## copy chart between PowerPoint presentations
```csharp
//Get the chart that is going to be copied.
IChart chart = presentation1.Slides[0].Shapes[0] as IChart;

//Copy chart from the first document to the second document.
presentation2.Slides.Append();
presentation2.Slides[1].Shapes.CreateChart(chart, new RectangleF(100, 100, 500, 300), -1);
```

---

# spire.presentation csharp chart
## copy chart within same presentation
```csharp
//Get the chart that is going to be copied.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Copy the chart from the first slide to the specified location of the second slide within the same document.
ISlide slide1 = presentation.Slides.Append();
slide1.Shapes.CreateChart(chart, new RectangleF(100, 100, 500, 300), 0);
```

---

# spire.presentation csharp chart
## create 100 percent stacked bar chart
```csharp
//Create a PowerPoint document
Presentation presentation = new Presentation();

//Set slide size
presentation.SlideSize.Type = SlideSizeType.Screen16x9;
SizeF slidesize = presentation.SlideSize.Size;

var slide = presentation.Slides[0];

//Append a 100% stacked bar chart
RectangleF rect = new RectangleF(20, 20, slidesize.Width - 40, slidesize.Height - 40);
IChart chart = slide.Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Bar100PercentStacked, rect);

//Set up chart data with labels and values
String[] columnlabels = { "Series 1", "Series 2", "Series 3" };
string[] rowlabels = { "Category 1", "Category 2", "Category 3" };
double[,] values = new double[3, 3] { { 20.83233, 10.34323, -10.354667 }, { 10.23456, -12.23456, 23.34456 }, { 12.34345, -23.34343, -13.23232 } };

//Insert the column labels
for (Int32 c = 0; c < columnlabels.Length; ++c)
    chart.ChartData[0, c + 1].Text = columnlabels[c];

//Insert the row labels
for (Int32 r = 0; r < rowlabels.Length; ++r)
    chart.ChartData[r + 1, 0].Text = rowlabels[r];

//Insert the values
for (Int32 r = 0; r < rowlabels.Length; ++r)
{
    for (Int32 c = 0; c < columnlabels.Length; ++c)
    {
        chart.ChartData[r + 1, c + 1].Value = Math.Round(values[r, c], 2);
    }
}

chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, columnlabels.Length];
chart.Categories.CategoryLabels = chart.ChartData[1, 0, rowlabels.Length, 0];

//Set the position of category axis
chart.PrimaryCategoryAxis.Position = AxisPositionType.Left;
chart.SecondaryCategoryAxis.Position = AxisPositionType.Left;
chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow;

//Set the data, font and format for the series of each column
for (Int32 c = 0; c < columnlabels.Length; ++c)
{
    chart.Series[c].Values = chart.ChartData[1, c + 1, rowlabels.Length, c + 1];
    chart.Series[c].Fill.FillType = FillFormatType.Solid;
    chart.Series[c].InvertIfNegative = false;

    for (Int32 r = 0; r < rowlabels.Length; ++r)
    {
        var label = chart.Series[c].DataLabels.Add();
        label.LabelValueVisible = true;
        chart.Series[c].DataLabels[r].HasDataSource = false;
        chart.Series[c].DataLabels[r].NumberFormat = "0#\\%";
        chart.Series[c].DataLabels.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 12;
    }
}

//Set the color of the Series
chart.Series[0].Fill.SolidColor.Color = Color.YellowGreen;
chart.Series[1].Fill.SolidColor.Color = Color.Red;
chart.Series[2].Fill.SolidColor.Color = Color.Green;

TextFont font = new TextFont("Tw Cen MT");

//Set the font and size for chart legend
for (int k = 0; k < chart.ChartLegend.EntryTextProperties.Length; k++)
{
    chart.ChartLegend.EntryTextProperties[k].LatinFont = font;
    chart.ChartLegend.EntryTextProperties[k].FontHeight = 20;
}
```

---

# spire.presentation csharp chart
## create Box and Whisker chart in PowerPoint presentation
```csharp
// Create a PPT document
Presentation ppt = new Presentation();

// Insert a BoxAndWhisker chart to the first slide 
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.BoxAndWhisker, new RectangleF(50, 50, 500, 400), false);

// Set chart data, series and categories
chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 3];
chart.Categories.CategoryLabels = chart.ChartData[1, 0, 18, 0];

chart.Series[0].Values = chart.ChartData[1, 1, 18, 1];
chart.Series[1].Values = chart.ChartData[1, 2, 18, 2];
chart.Series[2].Values = chart.ChartData[1, 3, 18, 3];

// Configure series properties specific to BoxAndWhisker chart
chart.Series[0].ShowInnerPoints = false;
chart.Series[0].ShowOutlierPoints = true;
chart.Series[0].ShowMeanMarkers = true;
chart.Series[0].ShowMeanLine = true;
chart.Series[0].QuartileCalculationType = QuartileCalculation.ExclusiveMedian;

chart.Series[1].ShowInnerPoints = false;
chart.Series[1].ShowOutlierPoints = true;
chart.Series[1].ShowMeanMarkers = true;
chart.Series[1].ShowMeanLine = true;
chart.Series[1].QuartileCalculationType = QuartileCalculation.InclusiveMedian;

chart.Series[2].ShowInnerPoints = false;
chart.Series[2].ShowOutlierPoints = true;
chart.Series[2].ShowMeanMarkers = true;
chart.Series[2].ShowMeanLine = true;
chart.Series[2].QuartileCalculationType = QuartileCalculation.ExclusiveMedian;

// Set chart title and legend
chart.HasLegend = true;
chart.ChartTitle.TextProperties.Text = "BoxAndWhisker";
chart.ChartLegend.Position = ChartLegendPositionType.Top;
```

---

# spire.presentation csharp bubble chart
## create bubble chart in PowerPoint presentation
```csharp
//Create a PPT file.
Presentation presentation = new Presentation();

//Add bubble chart
RectangleF rect1 = new RectangleF(90, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Bubble, rect1, false);

//Chart title
chart.ChartTitle.TextProperties.Text = "Bubble Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Attach the data to chart
Double[] xdata = new Double[] { 7.7, 8.9, 1.0, 2.4 };
Double[] ydata = new Double[] { 15.2, 5.3, 6.7, 8 };
Double[] size = new Double[] { 1.1, 2.4, 3.7, 4.8 };

chart.ChartData[0, 0].Text = "X-Value";
chart.ChartData[0, 1].Text = "Y-Value";
chart.ChartData[0, 2].Text = "Size";

for (Int32 i = 0; i < xdata.Length; ++i)
{
    chart.ChartData[i + 1, 0].Value = xdata[i];
    chart.ChartData[i + 1, 1].Value = ydata[i];
    chart.ChartData[i + 1, 2].Value = size[i];
}

//Set series label
chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];

chart.Series[0].XValues = chart.ChartData["A2", "A5"];
chart.Series[0].YValues = chart.ChartData["B2", "B5"];
chart.Series[0].Bubbles.Add(chart.ChartData["C2"]);
chart.Series[0].Bubbles.Add(chart.ChartData["C3"]);
chart.Series[0].Bubbles.Add(chart.ChartData["C4"]);
chart.Series[0].Bubbles.Add(chart.ChartData["C5"]);
```

---

# Spire.Presentation C# Chart
## Create Clustered Column Chart in PowerPoint
```csharp
//Create a PPT file
Presentation presentation = new Presentation();

//Add clustered column chart
RectangleF rect1 = new RectangleF(90, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect1, false);

//Chart title
chart.ChartTitle.TextProperties.Text = "Clustered Column Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Set series text
chart.ChartData[0, 1].Text = "Series1";
chart.ChartData[0, 2].Text = "Series2";

//Set category text
chart.ChartData[1, 0].Text = "Category 1";
chart.ChartData[2, 0].Text = "Category 2";
chart.ChartData[3, 0].Text = "Category 3";
chart.ChartData[4, 0].Text = "Category 4";
 
//Set series label
chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];
//Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];

//Set values for series
chart.Series[0].Values = chart.ChartData["B2", "B5"];
chart.Series[1].Values = chart.ChartData["C2", "C5"];
```

---

# Spire.Presentation C# Combination Chart
## Create a combination chart with column and line series in PowerPoint
```csharp
//Create a presentation instance
Presentation presentation = new Presentation();

//Insert a column clustered chart
RectangleF rect = new RectangleF(100, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect);

//Set chart title
chart.ChartTitle.TextProperties.Text = "Monthly Sales Report";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Create a datatable
DataTable dataTable = new DataTable();
dataTable.Columns.Add(new DataColumn("Month", Type.GetType("System.String")));
dataTable.Columns.Add(new DataColumn("Sales", Type.GetType("System.Int32")));
dataTable.Columns.Add(new DataColumn("Growth rate", Type.GetType("System.Decimal")));
dataTable.Rows.Add("January", 200, 0.6);
dataTable.Rows.Add("February", 250, 0.8);
dataTable.Rows.Add("March", 300, 0.6);
dataTable.Rows.Add("April", 150, 0.2);
dataTable.Rows.Add("May", 200, 0.5);
dataTable.Rows.Add("June", 400, 0.9);

//Import data from datatable to chart data
for (int c = 0; c < dataTable.Columns.Count; c++)
{
    chart.ChartData[0, c].Text = dataTable.Columns[c].Caption;
}
for (int r = 0; r < dataTable.Rows.Count; r++)
{
    object[] datas = dataTable.Rows[r].ItemArray;
    for (int c = 0; c < datas.Length; c++)
    {
        chart.ChartData[r + 1, c].Value = datas[c];
    }
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];

//Set categories labels    
chart.Categories.CategoryLabels = chart.ChartData["A2", "A7"];

//Assign data to series values
chart.Series[0].Values = chart.ChartData["B2", "B7"];
chart.Series[1].Values = chart.ChartData["C2", "C7"];

//Change the chart type of serie 2 to line with markers
chart.Series[1].Type = ChartType.LineMarkers;

//Plot data of series 2 on the secondary axis
chart.Series[1].UseSecondAxis = true;

//Set the number format as percentage 
chart.SecondaryValueAxis.NumberFormat = "0%";

//Hide gridlinkes of secondary axis
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;

//Set overlap
chart.OverLap = -50;

//Set gapwidth
chart.GapWidth = 200;
```

---

# spire.presentation csharp chart
## create 3D cylinder clustered chart in PowerPoint
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Insert chart
RectangleF rect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 200, 85, 400, 400);
IChart chart = presentation.Slides[0].Shapes.AppendChart(Spire.Presentation.Charts.ChartType.Cylinder3DClustered, rect);

//Add chart Title
chart.ChartTitle.TextProperties.Text = "Report";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Load data from datatable to chart
chart.Series.SeriesLabel = chart.ChartData["B1", "D1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A7"];
chart.Series[0].Values = chart.ChartData["B2", "B7"];
chart.Series[0].Fill.FillType = FillFormatType.Solid;
chart.Series[0].Fill.SolidColor.KnownColor = KnownColors.Brown;
chart.Series[1].Values = chart.ChartData["C2", "C7"];
chart.Series[1].Fill.FillType = FillFormatType.Solid;
chart.Series[1].Fill.SolidColor.KnownColor = KnownColors.Green;
chart.Series[2].Values = chart.ChartData["D2", "D7"];
chart.Series[2].Fill.FillType = FillFormatType.Solid;
chart.Series[2].Fill.SolidColor.KnownColor = KnownColors.Orange;

//Set the 3D rotation
chart.RotationThreeD.XDegree = 10;
chart.RotationThreeD.YDegree = 10;
```

---

# spire.presentation csharp chart
## create doughnut chart in presentation
```csharp
//Create a ppt document
Presentation presentation = new Presentation();
RectangleF rect = new RectangleF(80, 100, 550, 320);

//Set background image
RectangleF rect2 = new RectangleF(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect2);
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.FloralWhite;

//Add a Doughnut chart
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Doughnut, rect, false);
chart.ChartTitle.TextProperties.Text = "Market share by country";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;

string[] countries = new string[] { "Guba", "Mexico", "France", "German" };
int[] sales = new int[] { 1800, 3000, 5100, 6200 };
chart.ChartData[0, 0].Text = "Countries";
chart.ChartData[0, 1].Text = "Sales";
for (int i = 0; i < countries.Length; ++i)
{
    chart.ChartData[i + 1, 0].Value = countries[i];
    chart.ChartData[i + 1, 1].Value = sales[i];
}
chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];

for (int i = 0; i < chart.Series[0].Values.Count; i++)
{
    ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
    cdp.Index = i;
    chart.Series[0].DataPoints.Add(cdp);
}
//Set the series color
chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.LightBlue;
chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.MediumPurple;
chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.DarkGray;
chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.DarkOrange;

chart.Series[0].DataLabels.LabelValueVisible = true;
chart.Series[0].DataLabels.PercentValueVisible = true;
chart.Series[0].DoughnutHoleSize = 60;
```

---

# spire.presentation csharp funnel chart
## create a funnel chart in PowerPoint presentation
```csharp
//Create PPT document
Presentation ppt = new Presentation();

//Create a Funnel chart to the first slide
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Funnel, new RectangleF(50, 50, 550, 400), false);

//Set series text
chart.ChartData[0, 1].Text = "Series 1";

//Set category text
string[] categories = { "Website Visits", "Download", "Uploads", "Requested price", "Invoice sent", "Finalized" };
for (int i = 0; i < categories.Length; i++)
{
    chart.ChartData[i + 1, 0].Text = categories[i];
}

//Fill data for chart
double[] values = { 50000, 47000, 30000, 15000, 9000, 5600 };
for (int i = 0; i < values.Length; i++)
{
    chart.ChartData[i + 1, 1].NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];

//Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1, 0, categories.Length, 0];

//Assign data to series values
chart.Series[0].Values = chart.ChartData[1, 1, values.Length, 1];

//Set the chart title
chart.ChartTitle.TextProperties.Text = "Funnel";
```

---

# Spire.Presentation C# Chart
## Create Histogram Chart in PowerPoint
```csharp
//Create PPT document
Presentation ppt = new Presentation();

//Add a Histogram chart
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Histogram, new RectangleF(50, 50, 500, 400), false);

//Set series text
chart.ChartData[0, 0].Text = "Series 1";

//Set series label
chart.Series.SeriesLabel = chart.ChartData[0, 0, 0, 0];

chart.PrimaryCategoryAxis.NumberOfBins = 7;
chart.PrimaryCategoryAxis.GapWidth = 20;
//Chart title
chart.ChartTitle.TextProperties.Text = "Histogram";
chart.ChartLegend.Position = ChartLegendPositionType.Bottom;
```

---

# spire.presentation csharp chart
## create line markers chart
```csharp
//Create a PPT file
Presentation presentation = new Presentation();

//Add line markers chart
RectangleF rect1 = new RectangleF(90, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.LineMarkers, rect1, false);

//Chart title
chart.ChartTitle.TextProperties.Text = "Line Makers Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Data for series
Double[] Series1 = new Double[] { 7.7, 8.9, 1.0, 2.4 };
Double[] Series2 = new Double[] { 15.2, 5.3, 6.7, 8 };

//Set series text
chart.ChartData[0, 1].Text = "Series1";
chart.ChartData[0, 2].Text = "Series2";

//Set category text
chart.ChartData[1, 0].Text = "Category 1";
chart.ChartData[2, 0].Text = "Category 2";
chart.ChartData[3, 0].Text = "Category 3";
chart.ChartData[4, 0].Text = "Category 4";

//Fill data for chart
for (Int32 i = 0; i < Series1.Length; ++i)
{
    chart.ChartData[i + 1, 1].Value = Series1[i];
    chart.ChartData[i + 1, 2].Value = Series2[i];
}

//Set series label
chart.Series.SeriesLabel = chart.ChartData["B1", "C1"];
//Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];

//Set values for series
chart.Series[0].Values = chart.ChartData["B2", "B5"];
chart.Series[1].Values = chart.ChartData["C2", "C5"];
```

---

# Spire.Presentation C# Map Chart
## Create a map chart in PowerPoint presentation using Spire.Presentation library
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//Insert a Map chart to the first slide 
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Map, new RectangleF(50, 50, 450, 450), false);
chart.ChartData[0, 1].Text = "series";

//Define some data.
string[] countries = { "China", "Russia", "France", "Mexico", "United States", "India", "Australia" };
for (int i = 0; i < countries.Length; i++)
{
    chart.ChartData[i + 1, 0].Text = countries[i];
}
int[] values = { 32, 20, 23, 17, 18, 6, 11 };
for (int i = 0; i < values.Length; i++)
{
    chart.ChartData[i + 1, 1].NumberValue = values[i];
}
chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];
chart.Categories.CategoryLabels = chart.ChartData[1, 0, 7, 0];
chart.Series[0].Values = chart.ChartData[1, 1, 7, 1];
```

---

# Spire.Presentation C# Pareto Chart
## Creating a Pareto chart in PowerPoint presentation
```csharp
//Create a Pareto chart in first slide
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.Pareto, new RectangleF(50, 50, 500, 400), false);

//Set series text
chart.ChartData[0, 1].Text = "Series 1";

//Set up chart data and configure chart
chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];
chart.Categories.CategoryLabels = chart.ChartData[1, 0, 28, 0];
chart.Series[0].Values = chart.ChartData[1, 1, 28, 1];
chart.PrimaryCategoryAxis.IsBinningByCategory = true;
chart.Series[1].Line.FillFormat.FillType = FillFormatType.Solid;
chart.Series[1].Line.FillFormat.SolidFillColor.Color = Color.Red;
chart.ChartTitle.TextProperties.Text = "Pareto";
chart.HasLegend = true;
chart.ChartLegend.Position = ChartLegendPositionType.Bottom;
```

---

# spire.presentation csharp chart
## create pie chart in powerpoint presentation
```csharp
//Insert a Pie chart to the first slide and set the chart title.
RectangleF rect1 = new RectangleF(40, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Pie, rect1, false);
chart.ChartTitle.TextProperties.Text = "Sales by Quarter";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Define some data.
string[] quarters = new string[] { "1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr" };
int[] sales = new int[] { 210, 320, 180, 500 };

//Append data to ChartData, which represents a data table where the chart data is stored.
chart.ChartData[0, 0].Text = "Quarters";
chart.ChartData[0, 1].Text = "Sales";
for (int i = 0; i < quarters.Length; ++i)
{
    chart.ChartData[i + 1, 0].Value = quarters[i];
    chart.ChartData[i + 1, 1].Value = sales[i];
}

//Set category labels, series label and series data.
chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];
chart.Categories.CategoryLabels = chart.ChartData["A2", "A5"];
chart.Series[0].Values = chart.ChartData["B2", "B5"];

//Add data points to series and fill each data point with different color.
for (int i = 0; i < chart.Series[0].Values.Count; i++)
{
    ChartDataPoint cdp = new ChartDataPoint(chart.Series[0]);
    cdp.Index = i;
    chart.Series[0].DataPoints.Add(cdp);
}
chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.RosyBrown;
chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.LightBlue;
chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.LightPink;
chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid;
chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.MediumPurple;

//Set the data labels to display label value and percentage value.
chart.Series[0].DataLabels.LabelValueVisible = true;
chart.Series[0].DataLabels.PercentValueVisible = true;
```

---

# spire.presentation csharp chart
## create scatter chart with markers in presentation
```csharp
//Create a presentation
Presentation pres = new Presentation();

//Insert a chart and set chart title and chart type
RectangleF rect1 = new RectangleF(90, 100, 550, 320);
IChart chart = pres.Slides[0].Shapes.AppendChart(ChartType.ScatterMarkers, rect1, false);
chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Set chart data
Double[] xdata = new Double[] { 2.7, 8.9, 10.0, 12.4 };
Double[] ydata = new Double[] { 3.2, 15.3, 6.7, 8 };

chart.ChartData[0, 0].Text = "X-Value";
chart.ChartData[0, 1].Text = "Y-Value";

for (Int32 i = 0; i < xdata.Length; ++i)
{
    chart.ChartData[i + 1, 0].Value = xdata[i];
    chart.ChartData[i + 1, 1].Value = ydata[i];
}

//Set the series label
chart.Series.SeriesLabel = chart.ChartData["B1", "B1"];

//Assign data to X axis, Y axis and Bubbles
chart.Series[0].XValues = chart.ChartData["A2", "A5"];
chart.Series[0].YValues = chart.ChartData["B2", "B5"];
```

---

# spire.presentation csharp chart
## create SunBurst chart
```csharp
//Create PPT document
Presentation ppt = new Presentation();

//Create a SunBurst chart to the first slide
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.SunBurst, new RectangleF(50, 50, 500, 400), false);

//Set series text
chart.ChartData[0, 3].Text = "Series 1";

//Set category text
string[,] categories = {{"Branch 1","Stem 1","Leaf 1"},{"Branch 1","Stem 1","Leaf 2"},{"Branch 1","Stem 1", "Leaf 3"},
     {"Branch 1","Stem 2","Leaf 4"},{"Branch 1","Stem 2","Leaf 5"},{"Branch 1","Leaf 6",null},{"Branch 1","Leaf 7", null},
     {"Branch 2","Stem 3","Leaf 8"},{"Branch 2","Leaf 9",null},{"Branch 2","Stem 4","Leaf 10"},{"Branch 2","Stem 4","Leaf 11"},
     {"Branch 2","Stem 5","Leaf 12"},{"Branch 3","Stem 5","Leaf 13"},{"Branch 3","Stem 6","Leaf 14"},{"Branch 3","Leaf 15",null}};
for (int i = 0; i < 15; i++)
{
    for (int j = 0; j < 3; j++)
        chart.ChartData[i + 1, j].Value = categories[i, j];
}

//Fill data for chart
double[] values = { 17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51 };
for (int i = 0; i < values.Length; i++)
{
    chart.ChartData[i + 1, 3].NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData[0, 3, 0, 3];

//Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1, 0, values.Length, 2];

//Assign data to series values
chart.Series[0].Values = chart.ChartData[1, 3, values.Length, 3];

chart.Series[0].DataLabels.CategoryNameVisible = true;
chart.ChartTitle.TextProperties.Text = "SunBurst";
chart.HasLegend = true;
chart.ChartLegend.Position = ChartLegendPositionType.Top;
```

---

# spire.presentation csharp treemap chart
## create and configure TreeMap chart in PowerPoint presentation
```csharp
//Create a TreeMap chart to the first slide
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.TreeMap, new RectangleF(50, 50, 500, 400), false);

//Set series text
chart.ChartData[0, 3].Text = "Series 1";

//Set series labels
chart.Series.SeriesLabel = chart.ChartData[0, 3, 0, 3];

//Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1, 0, values.Length, 2];

//Assign data to series values
chart.Series[0].Values = chart.ChartData[1, 3, values.Length, 3];

//Configure TreeMap chart properties
chart.Series[0].DataLabels.CategoryNameVisible = true;
chart.Series[0].TreeMapLabelOption = TreeMapLabelOption.Banner;
chart.ChartTitle.TextProperties.Text = "TreeMap";
chart.HasLegend = true;
chart.ChartLegend.Position = ChartLegendPositionType.Top;
```

---

# spire.presentation csharp waterfall chart
## create WaterFall chart in PowerPoint
```csharp
//Create PPT document
Presentation ppt = new Presentation();

//Create a WaterFall chart to the first slide
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.WaterFall, new RectangleF(50, 50, 500, 400), false);

//Set series text
chart.ChartData[0, 1].Text = "Series 1";

//Set category text
string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7" };
for (int i = 0; i < categories.Length; i++)
{
    chart.ChartData[i + 1, 0].Text = categories[i];
}

//Fill data for chart
double[] values = { 100, 20, 50, -40, 130, -60, 70 };
for (int i = 0; i < values.Length; i++)
{
    chart.ChartData[i + 1, 1].NumberValue = values[i];
}

//Set series labels
chart.Series.SeriesLabel = chart.ChartData[0, 1, 0, 1];

//Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1, 0, categories.Length, 0];

//Assign data to series values
chart.Series[0].Values = chart.ChartData[1, 1, values.Length, 1];

//Operate the third datapoint of first series
ChartDataPoint chartDataPoint = new ChartDataPoint(chart.Series[0]);
chartDataPoint.Index = 2;
chartDataPoint.SetAsTotal = true;
chart.Series[0].DataPoints.Add(chartDataPoint);

//Operate the sixth datapoint of first series
ChartDataPoint chartDataPoint2 = new ChartDataPoint(chart.Series[0]);
chartDataPoint2.Index = 5;
chartDataPoint2.SetAsTotal = true;
chart.Series[0].DataPoints.Add(chartDataPoint2);
chart.Series[0].ShowConnectorLines = true;
chart.Series[0].DataLabels.LabelValueVisible = true;

chart.ChartLegend.Position = ChartLegendPositionType.Right;
chart.ChartTitle.TextProperties.Text = "WaterFall";
```

---

# spire.presentation csharp chart
## delete chart legend entries
```csharp
//Get the chart.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Delete the first and the second legend entries from the chart.
chart.ChartLegend.DeleteEntry(0);
chart.ChartLegend.DeleteEntry(1);
```

---

# Spire.Presentation C# Chart Detection
## Detect if a chart has SwitchRowAndColumn setting enabled
```csharp
//Get the chart from the presentation
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Detect whether the chart has "SwitchRowAndColumn" setting
bool result = chart.IsSwitchRowAndColumn();
```

---

# spire.presentation csharp doughnut chart
## set doughnut chart hole size
```csharp
//Get the chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set hole size
Chart.Series[0].DoughnutHoleSize = 55;
```

---

# spire.presentation csharp chart
## edit chart data in presentation
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Change the value of the second datapoint of the first series
Chart.Series[0].Values[1].Value = 6;
```

---

# Spire.Presentation C# Explode Pie Chart
## Set explosion distance for pie chart series
```csharp
//Get the chart that needs to set the point explosion.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

chart.Series[0].Distance = 15;
```

---

# spire.presentation csharp chart marker
## fill picture in chart marker
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Load image file in ppt
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.png");
IImageData IImage = ppt.Images.Append(image);

//Create a ChartDataPoint object and specify the index
ChartDataPoint dataPoint = new ChartDataPoint(Chart.Series[0]);
dataPoint.Index = 0;

//Fill picture in marker
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Picture;
dataPoint.MarkerFill.Fill.PictureFill.Picture.EmbedImage = IImage;

//Set marker size
dataPoint.MarkerSize = 20;

//Add the data point in series
Chart.Series[0].DataPoints.Add(dataPoint);
```

---

# Spire.Presentation C# Chart Data Labels
## Format chart data labels with custom text, position, font, and color
```csharp
//Get the chart
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Get the chart series
ChartSeriesFormatCollection sers = chart.Series;

//Initialize four instances of series label and set parameters of each label
ChartDataLabel cd1 = sers[0].DataLabels.Add();                   
cd1.PercentageVisible = true;
cd1.TextFrame.Text = "Custom Datalabel1";
cd1.TextFrame.TextRange.FontHeight = 12;
cd1.TextFrame.TextRange.LatinFont =new TextFont("Lucida Sans Unicode");
cd1.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
cd1.TextFrame.TextRange.Fill.SolidColor.Color= Color.Green;

ChartDataLabel cd2 = sers[0].DataLabels.Add();
cd2.Position = ChartDataLabelPosition.InsideEnd;
cd2.PercentageVisible = true;
cd2.TextFrame.Text = "Custom Datalabel2";
cd2.TextFrame.TextRange.FontHeight = 10;
cd2.TextFrame.TextRange.LatinFont = new TextFont("Arial");
cd2.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
cd2.TextFrame.TextRange.Fill.SolidColor.Color = Color.OrangeRed;

ChartDataLabel cd3 = sers[0].DataLabels.Add();
cd3.Position = ChartDataLabelPosition.Center;
cd3.PercentageVisible = true;
cd3.TextFrame.Text = "Custom Datalabel3";
cd3.TextFrame.TextRange.FontHeight = 14;
cd3.TextFrame.TextRange.LatinFont = new TextFont("Calibri");
cd3.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
cd3.TextFrame.TextRange.Fill.SolidColor.Color = Color.Blue;

ChartDataLabel cd4 = sers[0].DataLabels.Add();
cd4.Position = ChartDataLabelPosition.InsideBase;
cd4.PercentageVisible = true;
cd4.TextFrame.Text = "Custom Datalabel4";
cd4.TextFrame.TextRange.FontHeight = 12;
cd4.TextFrame.TextRange.LatinFont = new TextFont("Lucida Sans Unicode");
cd4.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
cd4.TextFrame.TextRange.Fill.SolidColor.Color = Color.OliveDrab;
```

---

# Spire.Presentation C# Chart Axis
## Get values and units from chart axis
```csharp
//Get chart on the first slide
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Get unit from primary category axis
float majorUnit = chart.PrimaryCategoryAxis.MajorUnit;
ChartBaseUnitType unitType = chart.PrimaryCategoryAxis.MajorUnitScale;

//Get values from primary value axis
float minValue = chart.PrimaryValueAxis.MinValue;
float maxValue = chart.PrimaryValueAxis.MaxValue;
```

---

# Spire.Presentation C# Chart Axis Labels
## Group two-level axis labels in a PowerPoint chart
```csharp
//Get the chart from the slide
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Get the category axis from the chart
IChartAxis chartAxis = chart.PrimaryCategoryAxis;

//Group the axis labels that have the same first-level label
if (chartAxis.HasMultiLvlLbl)
{
    chartAxis.IsMergeSameLabel = true;
}
```

---

# spire.presentation csharp chart
## hide chart axis and gridlines
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Hide axis
Chart.PrimaryCategoryAxis.IsVisible = false;
Chart.PrimaryValueAxis.IsVisible = false;

//Remove gridline
Chart.PrimaryValueAxis.MajorGridTextLines.FillType = FillFormatType.None;
```

---

# spire.presentation csharp chart
## hide or show a series in a chart
```csharp
//Get the first slide.
ISlide slide = presentation.Slides[0];

//Get the first chart.
IChart chart = slide.Shapes[0] as IChart;

//Hide the first series of the chart.
chart.Series[0].IsHidden = true;

//Show the first series of the chart.
//chart.Series[0].IsHidden = false;
```

---

# spire.presentation csharp chart
## invert if negative for chart series
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set invert if negative
Chart.Series[0].InvertIfNegative = true;
```

---

# spire.presentation csharp chart axis
## modify chart category axis
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Modify the major unit
Chart.PrimaryCategoryAxis.IsAutoMajor = false;
Chart.PrimaryCategoryAxis.MajorUnit = 1;
Chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months;
```

---

# spire.presentation csharp chart
## create multiple category chart
```csharp
//Create a PPT file
Presentation presentation = new Presentation();

//Add line markers chart
RectangleF rect1 = new RectangleF(90, 100, 550, 320);
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect1, false);

//Chart title
chart.ChartTitle.TextProperties.Text = "Muli-Category";
chart.ChartTitle.TextProperties.IsCentered = true;
chart.ChartTitle.Height = 30;
chart.HasTitle = true;

//Data for series
Double[] Series1 = new Double[] { 7.7, 8.9, 7, 6,7, 8 };

//Set series text
chart.ChartData[0, 2].Text = "Series1";

//Set category text
chart.ChartData[1, 0].Text = "Grp 1";
chart.ChartData[3, 0].Text = "Grp 2";
chart.ChartData[5, 0].Text = "Grp 3";           

chart.ChartData[1, 1].Text = "A";
chart.ChartData[2, 1].Text = "B";
chart.ChartData[3, 1].Text = "C";
chart.ChartData[4, 1].Text = "D";
chart.ChartData[5, 1].Text = "E";
chart.ChartData[6, 1].Text = "F";

//Fill data for chart
for (int i = 0; i < Series1.Length; ++i)
{
    chart.ChartData[i + 1, 2].Value = Series1[i];
}

//Set series label
chart.Series.SeriesLabel = chart.ChartData["C1", "C1"];
//Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2", "B7"];

//Set values for series
chart.Series[0].Values = chart.ChartData["C2", "C7"];

//Set if the category axis has multiple levels
chart.PrimaryCategoryAxis.HasMultiLvlLbl = true;
//Merge same label
chart.PrimaryCategoryAxis.IsMergeSameLabel = true;
```

---

# spire.presentation csharp chart protection
## Protect chart data in PowerPoint presentation
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.LoadFromFile("Template_Ppt_2.pptx");

//Get the first shape from slide and convert it as IChart.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Set the Boolean value of IChart.IsDataProtect as true.
chart.IsDataProtect = true;

//Save to file.
presentation.SaveToFile("Result-ProtectChart.pptx", FileFormat.Pptx2013);
```

---

# spire.presentation csharp chart
## get chart data range
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//Load PPT file 
ppt.LoadFromFile("ChartSample2.pptx");

//Get chart on the first slide
IChart chart = ppt.Slides[0].Shapes[0] as IChart;
if (chart != null)
{
    int lastRow = chart.ChartData.LastRowIndex;
    int lastCol = chart.ChartData.LastColIndex;
    // Process the chart data range information
}
```

---

# spire.presentation csharp chart
## remove chart from PowerPoint slide
```csharp
//Get the first slide from the document.
ISlide slide = presentation.Slides[0];

//Remove chart from the slide.
for (int i = 0; i < slide.Shapes.Count; i++)
{
    IShape shape = slide.Shapes[i] as IShape;
    if (shape is IChart)
    {
        slide.Shapes.Remove(shape);
    }
}
```

---

# spire.presentation csharp chart axis
## remove tick marks from chart axis and set number format
```csharp
//Get the chart that need to be adjusted the number format and remove the tick marks.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Set percentage number format for the axis value of chart.
chart.PrimaryValueAxis.NumberFormat = "0#\\%";

//Remove the tick marks for value axis and category axis.
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkNone;
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkNone;
chart.PrimaryCategoryAxis.MajorTickMark = TickMarkType.TickMarkNone;
chart.PrimaryCategoryAxis.MinorTickMark = TickMarkType.TickMarkNone;
```

---

# spire.presentation csharp chart
## save chart as image
```csharp
//Create a PPT document 
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile("SaveChartAsImage.pptx");

//Save chart as image in .png format
Image image = presentation.Slides[0].Shapes.SaveAsImage(0);
image.Save("Chart_result.png", System.Drawing.Imaging.ImageFormat.Png);
```

---

# Spire.Presentation Bubble Chart Scaling
## Scale bubble chart size in PowerPoint presentation
```csharp
//Get the chart from the first presentation slide
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Scale the bubble size, the range value is from 0 to 300
chart.BubbleScale = 50;
```

---

# spire.presentation csharp chart axis
## set chart axis position
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set axis position
Chart.PrimaryValueAxis.CrossBetweenType = CrossBetweenType.MidpointOfCategory;
```

---

# spire.presentation csharp chart
## set chart axis type to date axis
```csharp
//Get the chart
IChart chart = presentation.Slides[0].Shapes[1] as IChart;

chart.PrimaryCategoryAxis.AxisType = Spire.Presentation.Charts.AxisType.DateAxis;
chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months;
```

---

# spire.presentation csharp chart border
## set chart border style in presentation
```csharp
//Get chart on the first slide
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Set border style
chart.Line.FillFormat.FillType = FillFormatType.Solid;
chart.Line.FillFormat.SolidFillColor.Color = Color.Red;
chart.BorderRoundedCorners = true;
```

---

# Spire.Presentation C# Chart Data Labels
## Set data label ranges for a chart in PowerPoint
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Add a ColumnStacked chart
IChart chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnStacked, new RectangleF(100, 100, 500, 400));

//Set data for the chart
CellRange cellRange = chart.ChartData["F1"];
cellRange.Text = "labelA";
cellRange = chart.ChartData["F2"];
cellRange.Text = "labelB";
cellRange = chart.ChartData["F3"];
cellRange.Text = "labelC";
cellRange = chart.ChartData["F4"];
cellRange.Text = "labelD";

//Set data label ranges
chart.Series[0].DataLabelRanges = chart.ChartData["F1", "F4"];

//Add data label
ChartDataLabel dataLabel1 = chart.Series[0].DataLabels.Add();
dataLabel1.ID = 0;
//Show the value
dataLabel1.LabelValueVisible = false;
//Show the label string
dataLabel1.ShowDataLabelsRange = true;
```

---

# spire.presentation csharp chart
## set chart data number format
```csharp
// Get chart on the first slide
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

// Set the number format for Axis
chart.PrimaryValueAxis.NumberFormat = "#,##0.00";

// Set the DataLabels format for Axis
chart.Series[0].DataLabels.LabelValueVisible = true;
chart.Series[0].DataLabels.PercentValueVisible = false;
chart.Series[0].DataLabels.NumberFormat = "#,##0.00";
chart.Series[0].DataLabels.HasDataSource = false;

// Set the number format for ChartData
for (int i = 1; i <= chart.Series[0].Values.Count; i++)
{
    chart.ChartData[i, 1].NumberFormat = "#,##0.00";
}
```

---

# spire.presentation csharp trendline
## Set color and name for trendline in PowerPoint chart
```csharp
//Find the first chart in the first Slide
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Find the first trendline in the chart
ITrendlines trendline = chart.Series[0].TrendLines[0] as ITrendlines;

//Set name for trendline
trendline.Name = "trendlineName";

//Set color for trendline
trendline.Line.FillType = FillFormatType.Solid;
trendline.Line.SolidFillColor.Color = Color.Red;
```

---

# Spire.Presentation C# Chart Data Label
## Set position of data label in a chart
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Add data label
ChartDataLabel label = Chart.Series[0].DataLabels.Add();
//Set the position of the label
label.X = 0.1f;
label.Y = 0.1f;
```

---

# spire.presentation csharp chart datapoint color
## Set colors for data points in a chart
```csharp
//Get the chart
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Initialize an instances of dataPoint
ChartDataPoint cdp1 = new ChartDataPoint(chart.Series[0]);

//Specify the datapoint order
cdp1.Index = 0;

//Set the color of the datapoint
cdp1.Fill.FillType = FillFormatType.Solid;
cdp1.Fill.SolidColor.KnownColor = KnownColors.Orange;

//Add the dataPoint to first series
chart.Series[0].DataPoints.Add(cdp1);

//Set the color for the other three data points
ChartDataPoint cdp2 = new ChartDataPoint(chart.Series[0]);
cdp2.Index = 1;
cdp2.Fill.FillType = FillFormatType.Solid;
cdp2.Fill.SolidColor.KnownColor = KnownColors.Gold;
chart.Series[0].DataPoints.Add(cdp2);

ChartDataPoint cdp3 = new ChartDataPoint(chart.Series[0]);
cdp3.Index = 2;
cdp3.Fill.FillType = FillFormatType.Solid;
cdp3.Fill.SolidColor.KnownColor = KnownColors.MediumPurple;
chart.Series[0].DataPoints.Add(cdp3);

ChartDataPoint cdp4 = new ChartDataPoint(chart.Series[0]);
cdp4.Index = 1;
cdp4.Fill.FillType = FillFormatType.Solid;
cdp4.Fill.SolidColor.KnownColor = KnownColors.ForestGreen;
chart.Series[0].DataPoints.Add(cdp4);
```

---

# spire.presentation csharp chart
## set display unit for chart axis
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set the display unit
Chart.PrimaryValueAxis.DisplayUnit = ChartDisplayUnitType.Hundreds;
```

---

# spire.presentation csharp chart
## Set distance from axis
```csharp
//Create a ppt document
Presentation ppt = new Presentation();

//Append ColumnClustered chart
IChart chart = ppt.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, new RectangleF(50, 50, 400, 400));

//Get the PrimaryCategory axis
IChartAxis chartAxis = chart.PrimaryCategoryAxis;

//Set "Distance from axis"
chartAxis.LabelsDistance = 200;
```

---

# Spire.Presentation Chart Gap Width
## Set gap width for chart in PowerPoint presentation
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set gap width
Chart.GapWidth = 50;
```

---

# spire.presentation csharp chart legend
## Set chart legend position and size
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set the legend positon
Chart.ChartLegend.Left = 20;
Chart.ChartLegend.Top = 20;

//Set the legend size
Chart.ChartLegend.Width = 250;
Chart.ChartLegend.Height = 30;
```

---

# spire.presentation csharp axis formatting
## Set number format for chart axis
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set the number format
Chart.PrimaryCategoryAxis.NumberFormat = "yyyy";
```

---

# spire.presentation csharp chart
## set percentage for chart labels
```csharp
float dataPontPercent = 0f;

for (int i = 0; i < Chart.Series.Count; i++)
{
    ChartSeriesDataFormat series = Chart.Series[i];
    //Get the total number
    float total = GetTotal(series.Values);
    for (int j = 0; j < series.Values.Count; j++)
    {
        //Get the percent
        dataPontPercent = float.Parse(series.Values[j].Text) / total * 100;
        //Add datalabels
        ChartDataLabel label = series.DataLabels.Add();
        label.LabelValueVisible = true;
        //Set the percent text for the label
        label.TextFrame.Paragraphs[0].Text = String.Format("{0:F2} %", dataPontPercent);
        label.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 12;
    }
}

private float GetTotal(CellRanges ranges)
{
    float total = 0;
    for (int i = 0; i < ranges.Count; i++)
    {
        total += float.Parse(ranges[i].Text);
    }

    return total;
}
```

---

# spire.presentation csharp chart data labels
## set position and style of chart data labels
```csharp
//Get the chart.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Add data label to chart and set its id.
ChartDataLabel label1 = chart.Series[0].DataLabels.Add();
label1.ID = 0;

//Set the default position of data label. This position is relative to the data markers.
//label1.Position = ChartDataLabelPosition.OutsideEnd;

//Set custom position of data label. This position is relative to the default position.
label1.X = 0.1f;
label1.Y = -0.1f;

//Set label value visible
label1.LabelValueVisible = true;

//Set legend key invisible
label1.LegendKeyVisible = false;

//Set category name invisible
label1.CategoryNameVisible = false;

//Set series name invisible
label1.SeriesNameVisible = false;

//Set Percentage invisible
label1.PercentageVisible = false;

//Set border style and fill style of data label
label1.Line.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
label1.Line.SolidFillColor.Color = Color.Blue;
label1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
label1.Fill.SolidColor.Color = Color.Orange;
```

---

# spire.presentation csharp map chart
## Set projection type of map chart
```csharp
// Get the chart
IChart chart = ppt.Slides[0].Shapes[9] as IChart;

// Get the type of projection
ProjectionType type = chart.Series[0].ProjectionType;

// Change the type of projection
chart.Series[0].ProjectionType = ProjectionType.Robinson;
```

---

# spire.presentation csharp chart title
## set rotation angle for chart title
```csharp
//Get the chart
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Set rotation angle for chart title
chart.ChartTitle.TextProperties.RotationAngle = -30;
```

---

# spire.presentation csharp chart
## set rotation angle for data labels
```csharp
//Get chart on the first slide
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Set the rotation angle for the datalabels of first serie
for (int i = 0; i < chart.Series[0].Values.Count; i++)
{
    ChartDataLabel datalabel = chart.Series[0].DataLabels.Add();
    datalabel.ID = i;
    datalabel.RotationAngle = 45;
}
```

---

# spire.presentation csharp chart
## Set rotation angle for value axis text in PowerPoint chart
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set the rotation angle for the text on the value axis
Chart.PrimaryValueAxis.TextRotationAngle = 45;
```

---

# spire.presentation csharp chart
## set series line color in PowerPoint chart
```csharp
//Get the first chart
IShape shape = ppt.Slides[0].Shapes[0];
if(shape is IChart)
{
    IChart chart = (IChart)shape;
    TextLineFormat seriesLine = chart.SeriesLine;
    seriesLine.FillType = FillFormatType.Solid;

    //Set the color of seriesLine
    seriesLine.FillFormat.SolidFillColor.Color = Color.Red;
}
```

---

# spire.presentation csharp chart series overlap
## Set the overlap percentage for chart series
```csharp
//Get chart from the presentation
IChart chart = ppt.Slides[0].Shapes[0] as IChart;

//Set series overlap
chart.OverLap = 50;
```

---

# spire.presentation csharp chart markers
## set size and style for chart markers in presentation
```csharp
//Get the chart from the presentation.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

for (int i = 0; i < chart.Series[0].Values.Count; i++)
{
    //Create a ChartDataPoint object and specify the index.
    ChartDataPoint dataPoint = new ChartDataPoint(chart.Series[0]);
    dataPoint.Index = i;

    //Set the fill color of the data marker.
    dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
    dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Yellow;

    //Set the line color of the data marker.
    dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
    dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.YellowGreen;

    //Set the size of the data marker.
    dataPoint.MarkerSize = 20;

    //Set the style of the data marker
    dataPoint.MarkerStyle = ChartMarkerType.Diamond;
    chart.Series[0].DataPoints.Add(dataPoint);
}
```

---

# spire.presentation csharp chart plot area
## Set size for chart plot area in PowerPoint presentation
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Set width and height for chart plot area
Chart.PlotArea.Width = 250;
Chart.PlotArea.Height = 300;
```

---

# spire.presentation csharp chart title font
## set text font for chart title in PowerPoint presentation
```csharp
//Get the chart.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Set the font for the text on chart title area.
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Arial Unicode MS");
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue;
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 50;
```

---

# spire.presentation csharp chart font
## Set text font for chart legend and axis
```csharp
//Get the chart.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Set the font for the text on Chart Legend area.
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Green;
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Arial Unicode MS");

//Set the font for the text on Chart Axis area.
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10;
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = new TextFont("Arial Unicode MS");
```

---

# spire.presentation csharp chart axis
## Set tick mark labels on category axis
```csharp
//Get the chart from the PowerPoint slide.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Rotate tick labels.
chart.PrimaryCategoryAxis.TextRotationAngle = 45;

//Specify interval between labels.
chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = false;
chart.PrimaryCategoryAxis.TickLabelSpacing = 2;

//Change position.
chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh;
```

---

# spire.presentation csharp chart
## set tick marks interval for chart axis
```csharp
// Get chart from presentation
IChart chart = ppt.Slides[0].Shapes[0] as IChart;
IChartAxis chartAxis = chart.PrimaryCategoryAxis;
// Set tick marks interval
chartAxis.TickMarkSpacing = 2;
```

---

# spire.presentation csharp chart labels
## show chart labels in presentation
```csharp
//Get chart on the first slide
IChart Chart = ppt.Slides[0].Shapes[0] as IChart;

//Show data labels
Chart.Series[0].DataLabels.LabelValueVisible = true;
Chart.Series[0].DataLabels.CategoryNameVisible = true;
Chart.Series[0].DataLabels.SeriesNameVisible = true;
```

---

# spire.presentation csharp chart
## vary colors of same series data markers
```csharp
//Get the chart from the presentation.
IChart chart = presentation.Slides[0].Shapes[0] as IChart;

//Create a ChartDataPoint object and specify the index.
ChartDataPoint dataPoint = new ChartDataPoint(chart.Series[0]);
dataPoint.Index = 0;

//Set the fill color of the data marker.
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Red;

//Set the line color of the data marker.
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Red;

//Add the data point to the point collection of a series.
chart.Series[0].DataPoints.Add(dataPoint);

dataPoint = new ChartDataPoint(chart.Series[0]);
dataPoint.Index = 1;
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Black;
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Black;
chart.Series[0].DataPoints.Add(dataPoint);

dataPoint = new ChartDataPoint(chart.Series[0]);
dataPoint.Index = 2;
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid;
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.Blue;
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid;
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.Blue;
chart.Series[0].DataPoints.Add(dataPoint);
```

---

# Spire.Presentation C# Conversion
## Convert ODP to PDF
```csharp
Presentation presentation = new Presentation();

//Load ODP file from disk
presentation.LoadFromFile("input.odp", FileFormat.ODP);

String result = "output.pdf";

//Save to file.
presentation.SaveToFile(result, FileFormat.PDF);
```

---

# Spire.Presentation C# PDF Conversion
## Convert presentation to PDF with default font
```csharp
// Create a presentation object
Presentation ppt = new Presentation();

// The font is preferred to convert to pdf or pictures, when the font used in the document is not installed in the system
Presentation.SetDefaultFontName("Arial");

// Save to PDF format
ppt.SaveToFile("ConvertPdfWithDefaultFont_out.pdf", FileFormat.PDF);
```

---

# spire.presentation pps to pptx conversion
## convert pps file to pptx format using spire.presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();
//Load file
ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Conversion.pps");

//Save the PPS document to PPTX file format
string result = "ConvertPPSToPPTX.pptx";
ppt.SaveToFile(result, FileFormat.Pptx2013);
```

---

# spire.presentation csharp conversion
## convert PowerPoint to OFD format
```csharp
// Create Presentation
Presentation presentation = new Presentation();

// Load ppt file
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\CopyParagraph.pptx");

// Save the PPT document to OFD format
String result = "ConvertPPTToOFD_result.ofd";
presentation.SaveToFile(result, Spire.Presentation.FileFormat.OFD);
```

---

# spire.presentation csharp pdf conversion
## convert unhidden slides to pdf format
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile("HideSlide1.pptx");

//Convert the PPT unhidden slides to PDF file format 
presentation.SaveToPdfOption.ContainHiddenSlides = false;
string result = "ToPdf.pdf";
presentation.SaveToFile(result, FileFormat.PDF);
```

---

# spire.presentation csharp conversion
## convert PowerPoint slide to TIFF with custom size
```csharp
//Create a new PPT document
Presentation newPresentation = new Presentation();

//Remove the default slide 
newPresentation.Slides.RemoveAt(0);

//Define a custom size
SizeF size = new SizeF(200F, 200F);

//Set PPT slide size
newPresentation.SlideSize.Size = size;

//Insert a slide (obtained from another presentation)
newPresentation.Slides.Insert(0, slide);

//Save as TIFF format
newPresentation.SaveToFile("output.tiff", Spire.Presentation.FileFormat.Tiff);
```

---

# Spire.Presentation C# Slide Conversion
## Convert individual PowerPoint slide to HTML format
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk.
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

//Get the first slide
ISlide slide = presentation.Slides[0];

//String for output file 
String result = "Output.html";

//Save the first slide to HTML 
slide.SaveToFile(result, Spire.Presentation.FileFormat.Html);
```

---

# spire.presentation csharp file conversion
## load and save DPS and DPT presentation files
```csharp
//Create Presentation
Presentation presentation = new Presentation();

//Load .dps or .dpt file
presentation.LoadFromFile("sample.dps", Spire.Presentation.FileFormat.Dps);
//presentation.LoadFromFile("sample.dpt", Spire.Presentation.FileFormat.Dpt);

//Save the .dps or .dpt file
String result = "LoadSaveDPSAndDPT_result.dps";
presentation.SaveToFile(result, Spire.Presentation.FileFormat.Dps);
//presentation.SaveToFile("LoadSaveDPSAndDPT_result.dpt", Spire.Presentation.FileFormat.Dpt);
```

---

# spire.presentation csharp conversion
## convert powerpoint slide to svg format
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\OneSlideToSVG.pptx");

//Convert the second slide to SVG
byte[] svgByte = presentation.Slides[1].SaveToSVG();            
File.WriteAllBytes("OneSlideToSVG.svg", svgByte);
```

---

# spire.presentation csharp svg conversion
## Convert PowerPoint shapes to SVG with underline decoration option
```csharp
// Save the underline as decoration when converting to Svg
ppt.SaveToSvgOption.SaveUnderlineAsDecoration = true;

// Save to Svg
byte[] svgByte = ppt.Slides[0].Shapes[0].SaveAsSvgInSlide();
```

---

# Spire.Presentation C# Set Global Custom Fonts
## Set custom fonts directory for presentation conversion
```csharp
//Set global custom fonts 
Presentation.SetCustomFontsDirctory(@"..\..\..\..\..\..\Data\fonts");

//Create a PPT document
Presentation ppt = new Presentation();

//Load PPT file 
ppt.LoadFromFile(@"..\..\..\..\..\..\Data\toPDF.pptx");

//Save the PPT to PDF file format
String result = "output.pdf";
ppt.SaveToFile(result, FileFormat.PDF);
```

---

# spire.presentation csharp slide conversion
## convert specific slide to PDF
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk.
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ChangeSlidePosition.pptx");

//Get the second slide
ISlide slide= presentation.Slides[1];

//String for output file 
String result = "Output.pdf";

//Save the second slide to PDF
slide.SaveToFile(result, Spire.Presentation.FileFormat.PDF);
```

---

# spire.presentation csharp emf conversion
## convert powerpoint slide to emf image
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile("presentation.pptx");

//Save to EMF image
presentation.Slides[0].SaveAsEMF("ToEMFImage.emf");
```

---

# spire.presentation csharp conversion
## convert presentation to html
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Load file
ppt.LoadFromFile("Conversion.pptx");

//Save the document to HTML format
string result = "ToHTML.html";
ppt.SaveToFile(result, FileFormat.Html);
```

---

# spire.presentation csharp conversion
## convert presentation slides to images
```csharp
//Save presentation slides to images
for (int i = 0; i < presentation.Slides.Count; i++)
{
    String fileName = String.Format("ToImage-img-{0}.png", i);
    Image image = presentation.Slides[i].SaveAsImage();
    image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
}
```

---

# spire.presentation csharp conversion
## convert presentation to markdown format
```csharp
// Create and load the file 
Presentation ppt = new Presentation();
ppt.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractText.pptx");
// Convert to markdown format
ppt.SaveToFile("ToMarkdown.md", FileFormat.Markdown);
ppt.Dispose();
```

---

# spire.presentation csharp pdf conversion
## convert powerpoint presentation to pdf
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.pptx");

//Save the PPT to PDF file format
presentation.SaveToFile("ToPdf.pdf", FileFormat.PDF);
```

---

# spire.presentation csharp pdf conversion
## convert powerpoint to pdf/a formats
```csharp
// Create PPT document and load file
Presentation ppt = new Presentation();
ppt.LoadFromFile("input.pptx");

// Save the PPT to PDF_A1A
ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1A;
ppt.SaveToFile("ToPDF_A1A.pdf", FileFormat.PDF);

// Save the PPT to PDF_A1B
ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
ppt.SaveToFile("ToPDF_A1B.pdf", FileFormat.PDF);

// Save the PPT to PDF_A2A
ppt.SaveToPdfOption.PdfConformanceLevel = PdfConformanceLevel.Pdf_A2A;
ppt.SaveToFile("ToPDF_A2A.pdf", FileFormat.PDF);
```

---

# spire.presentation csharp conversion
## Convert presentation to PDF with specific page size
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile("input.pptx");

//Set A4 page size
presentation.SlideSize.Type = SlideSizeType.A4;

//Set landscape orientation
presentation.SlideSize.Orientation = SlideOrienation.Landscape;

//Save the PPT to PDF file format
presentation.SaveToFile("result.pdf", FileFormat.PDF);
```

---

# spire.presentation csharp conversion
## convert PPT to PPTX format
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT file from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\ToPPTX.ppt");

//Save the PPT document to PPTX file format
presentation.SaveToFile("ToPPTX.pptx", FileFormat.Pptx2010);
```

---

# spire.presentation csharp conversion
## convert slide to specific size image
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Save the first slide to Image and set the image size to 600*400
Image img = ppt.Slides[0].SaveAsImage(600, 400);
```

---

# Spire.Presentation C# Conversion
## Convert PowerPoint presentation to SVG files
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Retain note when converting a PPT document to SVG files
presentation.IsNoteRetained = true;

Queue<byte[]> svgBytes = presentation.SaveToSVG();
int count = svgBytes.Count;
for (int i = 0; i < count; i++)
{
    byte[] bt = svgBytes.Dequeue();
    String fileName = String.Format("ToSVG-{0}.svg", i);
    FileStream fs = new FileStream(fileName, FileMode.Create);
    fs.Write(bt, 0, bt.Length);
}
```

---

# spire.presentation csharp tiff conversion
## convert powerpoint slides to tiff image
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile("presentation.pptx");
Image[] images = new Image[presentation.Slides.Count];

//Save PPT to images
for (int i = 0; i < presentation.Slides.Count; i++)
{
    images[i] = presentation.Slides[i].SaveAsImage();
}

//Make TIFF image using images
JoinTiffImages(images, "output.tiff", EncoderValue.CompressionLZW);

//Function to get specified ImageCodecInfo
private static ImageCodecInfo GetEncoderInfo(string mimeType)
{
    ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
    for (int j = 0; j < encoders.Length; j++)
    {
        if (encoders[j].MimeType == mimeType)
            return encoders[j];
    }

    throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
}

//Function to make TIFF using images
public static void JoinTiffImages(Image[] images, string outFile, EncoderValue compressEncoder)
{
    //Use the save encoder
    System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.SaveFlag;

    EncoderParameters ep = new EncoderParameters(2);
    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
    ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)compressEncoder);

    Image pages = null;
    int frame = 0;
    ImageCodecInfo info = GetEncoderInfo("image/tiff");

    foreach (Image img in images)
    {
        if (frame == 0)
        {
            pages = img;

            //Save the first frame
            pages.Save(outFile, info, ep);
        }
        else
        {
            //Save the intermediate frames
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

            pages.SaveAdd(img, ep);
        }

        if (frame == images.Length - 1)
        {
            //Flush and close
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);
            pages.SaveAdd(ep);
        }

        frame++;
    }
}
```

---

# Spire.Presentation C# Conversion
## Convert PowerPoint presentation to XPS format
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();
//Load file
ppt.LoadFromFile(@"..\..\..\..\..\..\Data\Conversion.pptx");

//Save the the XPS file
string result = "ToXPS.xps";
ppt.SaveToFile(result, FileFormat.XPS);
```

---

# spire.presentation csharp image manipulation
## resize images in presentation
```csharp
float scale=0.5f;
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IEmbedImage)
        {
            IEmbedImage image = shape as IEmbedImage;
            image.Width = image.Width * scale;
            image.Height = image.Height * scale;
        }
    }
}
```

---

# spire.presentation csharp image cropping
## crop image in presentation slide
```csharp
//Get the first shape in first slide
IShape shape = ppt.Slides[0].Shapes[0];

//If the shape is SlidePicture
if (shape is SlidePicture)
{
    SlidePicture slidePicture = (SlidePicture)shape;
    //Crop image
    slidePicture.Crop(slidePicture.Left + 50f, slidePicture.Top + 50f, 100f, 200f);
}
```

---

# Spire.Presentation C# Image Extraction
## Extract images from PowerPoint presentation
```csharp
// Load a PPT document
Presentation ppt = new Presentation();
ppt.LoadFromFile("ExtractImage.pptx");

for (int i = 0; i < ppt.Images.Count; i++)
{
    string ImageName = string.Format("Images-{0}.png", i);
    // Extract image
    Image image = ppt.Images[i].Image;
    image.Save(ImageName);
}
```

---

# spire.presentation csharp extract images
## Extract images from a specific slide in PowerPoint presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();
//Load presentation file
ppt.LoadFromFile("presentation.pptx");

//Get the pictures from the second slide and save them as image files
int i = 0;
//Traverse all shapes in the second slide
foreach (IShape s in ppt.Slides[1].Shapes)
{
    //Check if it's a SlidePicture object
    if (s is SlidePicture)
    {
        //Save the image
        SlidePicture ps = s as SlidePicture;
        ps.PictureFill.Picture.EmbedImage.Image.Save(string.Format("image_{0}.png", i));
        i++;
    }
    //Check if it's a PictureShape object
    if (s is PictureShape)
    {
        //Save the image
        PictureShape ps = s as PictureShape;
        ps.EmbedImage.Image.Save(string.Format("image_{0}.png", i));
        i++;
    }
}
```

---

# Spire.Presentation CSharp EMF Image Insertion
## Insert EMF image into PowerPoint slide
```csharp
// EMF file path
string ImageFile = @"..\..\..\..\..\..\Data\InsertEMF.emf";

// Define image size
Image img = Image.FromFile(ImageFile);
float width = img.Width / 1.5f;
float height = img.Height / 1.5f;
RectangleF rect = new RectangleF(100, 100, width, height);

// Append the EMF in slide
IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile, rect);
image.Line.FillType = FillFormatType.None;
```

---

# spire.presentation csharp image
## insert image into presentation slide
```csharp
//Insert image to PPT
string ImageFile2 = @"..\..\..\..\..\..\Data\InsertImage.png";
RectangleF rect1 = new RectangleF(presentation.SlideSize.Size.Width / 2 - 280, 140, 120, 120);
IEmbedImage image = presentation.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, ImageFile2, rect1);
image.Line.FillType = FillFormatType.None;
```

---

# spire.presentation csharp remove images
## remove all images from a slide in PowerPoint presentation
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
  
for (int i = slide.Shapes.Count-1; i >=0; i--)
{
    //It is the SlidePicture object
    if (slide.Shapes[i] is SlidePicture)
    {
        slide.Shapes.RemoveAt(i);
    }            
}
```

---

# spire.presentation csharp image formatting
## set image frame format in presentation
```csharp
//Set the formatting of the image frame
pptImage.Line.FillFormat.FillType = FillFormatType.Solid;
pptImage.Line.FillFormat.SolidFillColor.Color = Color.LightBlue;
pptImage.Line.Width = 5;
pptImage.Rotation = -45;
```

---

# spire.presentation csharp image transparency
## Set transparency for an image in a presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Add a shape
IAutoShape shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(0, 0, 100, 100));
shape.Line.FillType = FillFormatType.None;
//Fill shape with image
shape.Fill.FillType = FillFormatType.Picture;
shape.Fill.PictureFill.Picture.Url = "imagePath";
shape.Fill.PictureFill.FillType = PictureFillType.Stretch;
//Set transparency on image
shape.Fill.PictureFill.Picture.Transparency = 50;
```

---

# spire.presentation csharp image update
## update an image in a presentation slide
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Append a new image to replace an existing image
IImageData image = ppt.Images.Append(Image.FromFile("image_path"));

//Replace the image which title is "image1" with the new image
foreach (IShape shape in slide.Shapes)
{
    if (shape is SlidePicture)
    {
        if (shape.AlternativeTitle == "image1")
        {
            (shape as SlidePicture).PictureFill.Picture.EmbedImage = image;
        }
    }
}
```

---

# spire.presentation csharp table cell image
## add image to powerpoint table cell
```csharp
//Get the first shape
ITable table = ppt.Slides[0].Shapes[0] as ITable;

//Load the image and insert it into table cell
IImageData pptImg = ppt.Images.Append(Image.FromFile("image_path.png"));

table[1, 1].FillFormat.FillType = FillFormatType.Picture;
table[1, 1].FillFormat.PictureFill.Picture.EmbedImage = pptImg;
table[1, 1].FillFormat.PictureFill.FillType = PictureFillType.Stretch;
```

---

# spire.presentation csharp table
## add row to table in powerpoint presentation
```csharp
//Get the table within the PowerPoint document.
ITable table = presentation.Slides[0].Shapes[0] as ITable;

//Get the second row.
TableRow row = table.TableRows[1];

//Clone the row and add it to the end of table.
table.TableRows.Append(row);
int rowCount = table.TableRows.Count;

//Get the last row.
TableRow lastRow = table.TableRows[rowCount - 1];

//Set new data of the first cell of last row.
lastRow[0].TextFrame.Text = " The first added cell";

//Set new data of the second cell of last row.
lastRow[1].TextFrame.Text = " The second added cell";
```

---

# spire.presentation csharp table column adjustment
## adjust table column width based on text content
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Get the table from the first slide of the sample document.
ISlide slide = presentation.Slides[0];
ITable table = slide.Shapes[0] as ITable;

//Adjust the first column width of table by text width.
table.ColumnsList[0].AdjustColumnByTextWidth();
```

---

# spire.presentation csharp table
## clone rows and columns in a presentation table
```csharp
// Create a presentation and access first slide
Presentation presentation = new Presentation();
ISlide sld = presentation.Slides[0];

// Define columns with widths and rows with heights
double[] widths = { 110, 110, 110 };
double[] heights = { 50, 30, 30, 30, 30 };

// Add table shape to slide
ITable table = presentation.Slides[0].Shapes.AppendTable(presentation.SlideSize.Size.Width / 2 - 275, 90, widths, heights);

// Clone row 1 at end of table
table.TableRows.Append(table.TableRows[0]);

// Clone row 2 as the 4th row of table
table.TableRows.Insert(3, table.TableRows[1]);

// Clone column 1 at end of table
table.ColumnsList.Add(table.ColumnsList[0]);

// Clone the 2nd column at 4th column index
table.ColumnsList.Insert(3, table.ColumnsList[1]);
```

---

# spire.presentation csharp table
## create table in PowerPoint presentation
```csharp
// Define table dimensions
Double[] widths = new double[] { 100, 100, 150, 100, 100 };
Double[] heights = new double[] { 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15 };

// Add new table to presentation slide
ITable table = presentation.Slides[0].Shapes.AppendTable(presentation.SlideSize.Size.Width / 2 - 275, 90, widths, heights);

// Add data to table cells
for (int row = 0; row < 13; row++)
    for (int col = 0; col < 5; col++)
    {
        // Fill the table with data
        table[col, row].TextFrame.Text = "Data";
        
        // Set the font for table cells
        table[col, row].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Narrow");
    }

// Set the alignment of the first row to Center
for (int col = 0; col < 5; col++)
{
    table[col, 0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
}

// Apply style to the table
table.StylePreset = TableStylePreset.LightStyle3Accent1;
```

---

# Spire.Presentation C# Table Editing
## Edit table data and style in PowerPoint presentation
```csharp
//Store the data used in replacement in string array.
string[] str = new string[] { "Germany", "Berlin", "Europe", "0152458", "20860000" };

ITable table = null;

//Get the table in PowerPoint document.
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        //Change the style of table.
        table.StylePreset = TableStylePreset.LightStyle1Accent2;

        for (int i = 0; i < table.ColumnsList.Count; i++)
        {
            //Replace the data in cell.
            table[i, 2].TextFrame.Text = str[i];

            //Set the highlight color.
            table[i, 2].TextFrame.TextRange.HighlightColor.Color = Color.BlueViolet;
        }
    }
}
```

---

# spire.presentation csharp table formatting
## Fill all table cells with color in PowerPoint presentation
```csharp
//Fill the table cells with color.
ITable table = null;
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;
        foreach (TableRow row in table.TableRows)
        {
            foreach (Cell cell in row)
            {
                cell.FillFormat.FillType = FillFormatType.Solid;
                cell.FillFormat.SolidColor.Color = Color.Pink;
            }
        }
    }
}
```

---

# Spire Presentation Table Row Color Fill
## Fill a particular table row with color in PowerPoint presentation
```csharp
//Fill particular table row with color.
ITable table = null;
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;
        
        TableRow row = table.TableRows[1];
        foreach (Cell cell in row)
        {
            cell.FillFormat.FillType = FillFormatType.Solid;
            cell.FillFormat.SolidColor.Color = Color.Pink;
        }
    }
}
```

---

# spire.presentation csharp table cell
## get border color of table cell
```csharp
//Get the table in the first slide
ITable table = presentation.Slides[0].Shapes[0] as ITable;

//Get borders' color of the first cell
string leftBorderColor = table[0, 0].BorderLeftDisplayColor;
string topBorderColor = table[0, 0].BorderTopDisplayColor;
string rightBorderColor = table[0, 0].BorderRightDisplayColor;
string bottomBorderColor = table[0, 0].BorderBottomDisplayColor;

//Get display color of the first cell
string cellColor = table[0,0].DisplayColor;
```

---

# Spire.Presentation C# Table
## Identify merged cells in PowerPoint tables
```csharp
foreach (IShape shape in slide.Shapes)
{
    //Verify if it is table
    if (shape is ITable)
    {
        ITable table = (ITable)shape;
        for (int r = 0; r < table.TableRows.Count; r++)
        {
            for (int c = 0; c < table.ColumnsList.Count; c++)
            {
                // Get cell
                Cell currentCell = table.TableRows[r][c];
                //Identify if it is merged cell
                if (currentCell.RowSpan > 1 || currentCell.ColSpan > 1)
                {
                    // Cell is merged with RowSpan and ColSpan properties
                    // The original cell position is at FirstRowIndex and FirstColumnIndex
                }
            }
        }                  
    }
}
```

---

# spire.presentation csharp table aspect ratio
## lock aspect ratio for tables in presentation
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];
foreach (IShape shape in slide.Shapes)
{
    //Verify if it is table
    if (shape is ITable)
    {
        ITable table = (ITable)shape;
        //Lock aspect ratio
        table.ShapeLocking.AspectRatioProtection = true;
    }
}
```

---

# spire.presentation csharp table
## merge table cells in PowerPoint presentation
```csharp
// Iterate through shapes on the first slide
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        // Merge the second row and third row of the first column
        table.MergeCells(table[0, 1], table[0, 2], false);

        // Merge another set of cells
        table.MergeCells(table[3, 4], table[4, 4], true);
    }
}
```

---

# spire.presentation csharp table
## remove rows and columns from PowerPoint table
```csharp
//Get the table in PPT document
ITable table = null;
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        //Remove the second column
        table.ColumnsList.RemoveAt(1, false);

        //Remove the second row
        table.TableRows.RemoveAt(1, false);
    }
}
```

---

# spire.presentation csharp table
## remove table border style in PowerPoint presentation
```csharp
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is ITable)
        {
            foreach (TableRow row in (shape as ITable).TableRows)
            {
                foreach (Cell cell in row)
                {
                    cell.BorderTop.FillType = FillFormatType.None;
                    cell.BorderBottom.FillType = FillFormatType.None;
                    cell.BorderLeft.FillType = FillFormatType.None;
                    cell.BorderRight.FillType = FillFormatType.None;
                }
            }
        }
    }
}
```

---

# spire.presentation csharp table removal
## remove all tables from a powerpoint slide
```csharp
//Get the tables within the PPT document.
List<IShape> shape_tems = new List<IShape>();

foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        //Add new table to table list.
        shape_tems.Add(shape);
    }
}

//Remove all the tables form the first slide.
foreach (IShape shape in shape_tems)
{
    presentation.Slides[0].Shapes.Remove(shape);
}
```

---

# Spire.Presentation Table Alignment
## Set horizontal and vertical alignment for table cells in PowerPoint presentation
```csharp
ITable table = null;
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        //Horizontal Alignment
        //Set the horizontal alignment for the cells in first column 
        table[0, 1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
        table[0, 2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
        table[0, 3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right;
        table[0, 4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify;

        //Vertical Alignment
        //Set the vertical alignment for the cells in second column 
        table[1, 1].TextAnchorType = TextAnchorType.Top;
        table[1, 2].TextAnchorType = TextAnchorType.Center;
        table[1, 3].TextAnchorType = TextAnchorType.Bottom;
        table[1, 4].TextAnchorType = TextAnchorType.None;

        //Both orientations
        //Set the both horizontal and vertical alignment for the cells in the third column 
        table[2, 1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
        table[2, 1].TextAnchorType = TextAnchorType.Top;

        table[2, 2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right;
        table[2, 2].TextAnchorType = TextAnchorType.Center;

        table[2, 3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify;
        table[2, 3].TextAnchorType = TextAnchorType.Bottom;

        table[2, 4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;
        table[2, 4].TextAnchorType = TextAnchorType.Top;
    }
}
```

---

# Spire.Presentation C# Table Borders
## Set borders for an existing table in a PowerPoint presentation
```csharp
//Get the table from the first slide of the sample document.
ISlide slide = presentation.Slides[0];
ITable table = slide.Shapes[0] as ITable;

//Set the border type as Inside and the border color as blue.
table.SetTableBorder(TableBorderType.Inside, 1, Color.Blue);
```

---

# Spire.Presentation C# Table Borders
## Set borders for newly created tables in a presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Set the table width and height for each table cell.
double[] tableWidth = new double[] { 100, 100, 100, 100, 100 };
double[] tableHeight = new double[] { 20, 20 };

//Traverse all the border type of the table.
foreach (TableBorderType item in Enum.GetValues(typeof(TableBorderType)))
{
  //Add a table to the presentation slide with the setting width and height
    ITable itable = presentation.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight);

    //Add some text to the table cell.
    itable.TableRows[0][0].TextFrame.Text = "Row";
    itable.TableRows[1][0].TextFrame.Text = "Column";

    //Set the border type, border width and the border color for the table.
    itable.SetTableBorder(item, 1.5, Color.Red);
}
```

---

# spire.presentation csharp table header
## set first row as table header in presentation
```csharp
// Find the table in the first slide
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = shape as ITable;
    }
}
// Set the first row as header
table.FirstRow = true;
```

---

# spire.presentation table manipulation
## set row height and column width for table in powerpoint
```csharp
//Get the table
ITable table = null;
foreach (IShape shape in ppt.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        //Set the height for the rows
        table.TableRows[0].Height = 100;
        table.TableRows[1].Height = 80;
        table.TableRows[2].Height = 60;
        table.TableRows[3].Height = 40;
        table.TableRows[4].Height = 20;

        //Set the column width
        table.ColumnsList[0].Width = 60;
        table.ColumnsList[1].Width = 80;
        table.ColumnsList[2].Width = 120;
        table.ColumnsList[3].Width = 140;
        table.ColumnsList[4].Width = 160;
    }
}
```

---

# spire.presentation csharp table border style
## set table border style in powerpoint presentation
```csharp
//Find the table by looping through all the slides, and then set borders for it. 
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is ITable)
        {
            foreach (TableRow row in (shape as ITable).TableRows)
            {
                foreach (Cell cell in row)
                {
                    cell.BorderTop.FillType = FillFormatType.Solid;
                    cell.BorderBottom.FillType = FillFormatType.Solid;
                    cell.BorderLeft.FillType = FillFormatType.Solid;
                    cell.BorderRight.FillType = FillFormatType.Solid;
                }
            }
        }
    }
}
```

---

# Spire.Presentation Table Styling
## Set table style in PowerPoint presentation
```csharp
//Get the table
ITable table = null;
foreach (IShape shape in ppt.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        //Set the table style from TableStylePreset and apply it to selected table
        table.StylePreset = TableStylePreset.MediumStyle1Accent2;
    }
}
```

---

# spire.presentation csharp table text formatting
## set text format for table cells in PowerPoint presentation
```csharp
// Verify if it is table
if (shape is ITable)
{
    ITable table = (ITable)shape;

    Cell cell1 = table.TableRows[0][0];
    // Set table cell's text alignment type 
    cell1.TextAnchorType = TextAnchorType.Top;
    // Set italic style
    cell1.TextFrame.TextRange.Format.IsItalic = TriState.True;

    Cell cell2 = table.TableRows[1][0];
    // Set table cell's foreground color
    cell2.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    cell2.TextFrame.TextRange.Fill.SolidColor.Color = Color.Green;
    // Set table cell's background color
    cell2.FillFormat.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    cell2.FillFormat.SolidColor.Color = Color.LightGray;

    Cell cell3 = table.TableRows[2][2];
    // Set table cell's font and font size
    cell3.TextFrame.TextRange.FontHeight = 12;
    cell3.TextFrame.TextRange.LatinFont = new TextFont("Arial Black");
    cell3.TextFrame.TextRange.HighlightColor.Color = Color.YellowGreen;

    Cell cell4 = table.TableRows[2][1];
    // Set table cell's margin and borders
    cell4.MarginLeft = 20;
    cell4.MarginTop = 30;
    cell4.BorderTop.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    cell4.BorderTop.SolidFillColor.Color = Color.Red;
    cell4.BorderBottom.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    cell4.BorderBottom.SolidFillColor.Color = Color.Red;
    cell4.BorderLeft.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    cell4.BorderLeft.SolidFillColor.Color = Color.Red;
    cell4.BorderRight.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    cell4.BorderRight.SolidFillColor.Color = Color.Red;
}
```

---

# spire.presentation csharp table
## split specific table cell
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];

//Get the table
ITable table = slide.Shapes[0] as ITable;

//Split cell [1, 2] into 3 rows and 2 columns
table[1, 2].Split(3, 2);
```

---

# spire.presentation csharp table
## traverse through cells in PowerPoint table
```csharp
//Get the table.
ITable table = null;
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ITable)
    {
        table = (ITable)shape;

        //Traverse through the cells of table.
        foreach (TableRow row in table.TableRows)
        {
            foreach (Cell cell in row)
            {
                content.AppendLine(cell.TextFrame.Text);
            }
            content.AppendLine("\n");
        }
    }
}
```

---

# spire.presentation csharp hyperlink
## add hyperlink to image in powerpoint
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Get the first slide.
ISlide slide = presentation.Slides[0];

//Add image to slide.
RectangleF rect = new RectangleF(480, 350, 160, 160);
IEmbedImage image = slide.Shapes.AppendEmbedImage(ShapeType.Rectangle, imageFilePath, rect);

//Add hyperlink to the image.
ClickHyperlink hyperlink = new ClickHyperlink("https://www.e-iceblue.com");
image.Click = hyperlink;
```

---

# Spire.Presentation C# Hyperlink
## Add hyperlink to SmartArt nodes
```csharp
//Get the smartArt shape
ISmartArt sr = ppt.Slides[0].Shapes[0] as ISmartArt;
//Add hyperlinks to the nodes
ISmartArtNode node = sr.Nodes[0];
node.Click = new ClickHyperlink(ppt.Slides[1]);
node = sr.Nodes[1];
node.Click = new ClickHyperlink(ppt.Slides[2]);
node = sr.Nodes[2];
node.Click = new ClickHyperlink(ppt.Slides[3]);
```

---

# spire.presentation csharp hyperlink
## add hyperlink to text in PowerPoint presentation
```csharp
//Find the text we want to add link to it.
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;
TextParagraph tp = shape.TextFrame.TextRange.Paragraph;
string temp = tp.Text;

//Split the original text.
string textToLink = "Spire.Presentation";
string[] strSplit = temp.Split(new string[] { "Spire.Presentation" }, StringSplitOptions.None);

//Clear all text.
tp.TextRanges.Clear();

//Add new text.
TextRange tr = new TextRange(strSplit[0]);
tp.TextRanges.Append(tr);

//Add the hyperlink.
tr = new TextRange(textToLink);
tr.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
tp.TextRanges.Append(tr);
```

---

# spire.presentation csharp hyperlink
## change hyperlink color in PowerPoint
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];

//Get the theme of the slide
Theme theme = slide.Theme;

//Change the color of hyperlink to red
theme.ColorScheme.HyperlinkColor.Color = Color.Red;
```

---

# spire.presentation csharp get linked slide
## retrieve target slide from hyperlink in presentation
```csharp
//Create Presentation
Presentation presentation = new Presentation();

//Load ppt file
presentation.LoadFromFile("linkedSlide.pptx");

//Get the second slide
ISlide slide = presentation.Slides[1];

//Get the first shape of the second slide
IAutoShape shape = slide.Shapes[0] as IAutoShape;

//Get the linked slide index
if (shape.Click.ActionType == HyperlinkActionType.GotoSlide)
{
    ISlide targetSlide = shape.Click.TargetSlide;
    MessageBox.Show("Linked slide number = " + targetSlide.SlideNumber);
}
```

---

# Spire.Presentation C# Hyperlink Outline Style
## Create a hyperlink with custom outline style in a PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add new shape to PPT document
RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 255, 120, 400, 100);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
shape.Fill.FillType = FillFormatType.None;
shape.Line.FillType = FillFormatType.None;

//Add a paragraph with hyperlink
TextParagraph para1 = new TextParagraph();
TextRange tr1 = new TextRange("Click to know more about Spire.Presentation");
tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
para1.TextRanges.Append(tr1);

//Set the format of textrange
tr1.Format.FontHeight = 20f;
tr1.IsItalic = TriState.True;

//Set the outline format of textrange
tr1.TextLineFormat.FillFormat.FillType = FillFormatType.Solid;
tr1.TextLineFormat.FillFormat.SolidFillColor.Color = Color.LightSeaGreen;
tr1.TextLineFormat.JoinStyle = LineJoinType.Round;
tr1.TextLineFormat.Width = 2f;

//Add the paragraph to shape
shape.TextFrame.Paragraphs.Append(para1); 
shape.TextFrame.Paragraphs.Append(new TextParagraph());
```

---

# spire.presentation csharp hyperlinks
## add hyperlinks to text in powerpoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Add new shape to PPT document
RectangleF rec = new RectangleF(presentation.SlideSize.Size.Width / 2 - 255, 120, 500, 280);
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec);
shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
shape.Line.Width = 0;

//Add title paragraph
TextParagraph para1 = new TextParagraph();
TextRange tr = new TextRange("E-iceblue");            
tr.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
tr.Fill.SolidColor.Color = System.Drawing.Color.Blue;
para1.TextRanges.Append(tr);
para1.Alignment = TextAlignmentType.Center;
shape.TextFrame.Paragraphs.Append(para1);
shape.TextFrame.Paragraphs.Append(new TextParagraph());

//Add hyperlink paragraph
TextParagraph para2 = new TextParagraph();
TextRange tr1 = new TextRange("Click to know more about Spire.Presentation.");
tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html";
para2.TextRanges.Append(tr1);
shape.TextFrame.Paragraphs.Append(para2);
shape.TextFrame.Paragraphs.Append(new TextParagraph());

//Add more hyperlink paragraphs
TextParagraph para3 = new TextParagraph();
TextRange tr2 = new TextRange("Click to visit E-iceblue Home page.");
tr2.ClickAction.Address = "https://www.e-iceblue.com/";
para3.TextRanges.Append(tr2);
shape.TextFrame.Paragraphs.Append(para3);
shape.TextFrame.Paragraphs.Append(new TextParagraph());

TextParagraph para4 = new TextParagraph();
TextRange tr3 = new TextRange("Click to go to the forum to raise questions.");
tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html";
para4.TextRanges.Append(tr3);
shape.TextFrame.Paragraphs.Append(para4);
shape.TextFrame.Paragraphs.Append(new TextParagraph());

TextParagraph para5 = new TextParagraph();
TextRange tr4 = new TextRange("Click to contact our sales team via email.");
tr4.ClickAction.Address = "mailto:sales@e-iceblue.com";
para5.TextRanges.Append(tr4);
shape.TextFrame.Paragraphs.Append(para5);
shape.TextFrame.Paragraphs.Append(new TextParagraph());

TextParagraph para6 = new TextParagraph();
TextRange tr5 = new TextRange("Click to contact our support team via email.");
tr5.ClickAction.Address = "mailto:support@e-iceblue.com";
para6.TextRanges.Append(tr5);
shape.TextFrame.Paragraphs.Append(para6);

//Format text
foreach (TextParagraph para in shape.TextFrame.Paragraphs)
{
    if (!string.IsNullOrEmpty(para.Text))
    {
        para.TextRanges[0].LatinFont = new TextFont("Lucida Sans Unicode");
        para.TextRanges[0].FontHeight = 20;
    }
}
```

---

# spire.presentation csharp hyperlink
## create hyperlink to specific slide in presentation
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Append a slide to it.
presentation.Slides.Append();

//Add a shape to the second slide.
IAutoShape shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(10, 50, 200, 50));
shape.Fill.FillType = FillFormatType.None;
shape.Line.FillType = FillFormatType.None;
shape.TextFrame.Text = "Jump to the first slide";

//Create a hyperlink based on the shape and the text on it, linking to the first slide.
ClickHyperlink hyperlink = new ClickHyperlink(presentation.Slides[0]);
shape.Click = hyperlink;
shape.TextFrame.TextRange.ClickAction = hyperlink;
```

---

# Spire.Presentation C# Hyperlink
## Create hyperlink to last viewed slide in PowerPoint presentation
```csharp
//Create a PPT document
Presentation ppt = new Presentation();

//Get specified slide
ISlide slide = ppt.Slides[0];

//Draw a shape
IAutoShape autoShape = slide.Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(100, 100, 100, 100));

//Link to last viewed slide show
autoShape.Click = ClickHyperlink.LastVievedSlide;
```

---

# spire.presentation hyperlink modification
## modify hyperlink in powerpoint presentation
```csharp
//Find the hyperlinks you want to edit.
IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];

//Edit the link text and the target URL.
shape.TextFrame.TextRange.ClickAction.Address = "http://www.e-iceblue.com";
shape.TextFrame.TextRange.Text = "E-iceblue";
```

---

# spire.presentation csharp hyperlink
## remove hyperlink from PowerPoint presentation
```csharp
//Get the shape and its text with hyperlink
IAutoShape shape = presentation.Slides[0].Shapes[0] as IAutoShape;

//Set the ClickAction property into null to remove the hyperlink
shape.TextFrame.TextRange.ClickAction = null;
```

---

# spire.presentation csharp audio extraction
## extract audio from PowerPoint presentation
```csharp
//Load a PPT document
Presentation presentation = new Presentation();
presentation.LoadFromFile(loadPath);

foreach (Shape shape in presentation.Slides[0].Shapes)
{
    if (shape is IAudio)
    {
        IAudio audio = shape as IAudio;
        byte[] AudioData = audio.Data.Data;

        using (FileStream fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
        {
            fs.Write(AudioData, 0, AudioData.Length);
        }
    }
}
```

---

# Spire.Presentation Video Extraction
## Extract videos from PowerPoint presentations
```csharp
// Create PPT document
Presentation presentation = new Presentation();

// Load the PPT document
presentation.LoadFromFile("video.pptx");

// Define a counter for output files
int videoIndex = 0;

// Traverse all the slides of PPT file
foreach (ISlide slide in presentation.Slides)
{
    // Traverse all the shapes of slides
    foreach (IShape shape in slide.Shapes)
    {
        // If shape is IVideo
        if (shape is IVideo)
        {
            // Create a unique filename for each video
            string outputFileName = $"extracted_video_{videoIndex}.avi";
            
            // Save the video
            (shape as IVideo).EmbeddedVideoData.SaveToFile(outputFileName);
            
            // Increment the counter
            videoIndex++;
        }
    }
}
```

---

# spire.presentation csharp audio
## hide audio during presentation show
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];

//Hide Audio during show
foreach (Shape shape in slide.Shapes)
{
    if (shape is IAudio)
    {
        IAudio audio = shape as IAudio;
        audio.HideAtShowing = true;
    }
}
```

---

# Spire.Presentation C# Audio Insertion
## Insert audio into PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Insert audio into the document
RectangleF audioRect = new RectangleF(220, 240, 80, 80);
presentation.Slides[0].Shapes.AppendAudioMedia(Path.GetFullPath(@"..\..\..\..\..\..\Data\Music.wav"), audioRect);
```

---

# Spire.Presentation C# Video Insertion
## Insert video into PowerPoint presentation
```csharp
//Insert video into the document
RectangleF videoRect = new RectangleF(presentation.SlideSize.Size.Width / 2 - 125, 240, 150, 150);
IVideo video = presentation.Slides[0].Shapes.AppendVideoMedia(Path.GetFullPath(@"..\..\..\..\..\..\Data\Video.mp4"), videoRect);
video.PictureFill.Picture.Url = @"..\..\..\..\..\..\Data\Video.png";
```

---

# spire.presentation audio properties
## obtain sound effect properties from presentation
```csharp
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.Slides[0];

//Get the audio in a time node
TimeNodeAudio audio = slide.Timeline.MainSequence[0].TimeNodeAudios[0];

//Get the properties of the audio, such as sound name, volume or detect if it's mute
string soundName = audio.SoundName;
double volume = audio.Volume;
bool isMute = audio.IsMute;
```

---

# spire.presentation csharp video replacement
## Replace videos in PowerPoint presentation
```csharp
// Get the videos collection from the presentation
VideoCollection videos = presentation.Videos;

// Traverse all the slides of PPT file
foreach (ISlide slide in presentation.Slides)
{
    // Traverse all the shapes of slides
    foreach (Shape shape in slide.Shapes)
    {
        // If shape is IVideo
        if (shape is IVideo)
        {
            // Replace the video
            IVideo video = shape as IVideo;
            // Create video data from byte array
            VideoData videoData = videos.Append(newVideoBytes);
            video.EmbeddedVideoData = videoData;
        }
    }
}
```

---

# spire.presentation video play mode
## Set play mode for videos in PowerPoint presentation
```csharp
//Find the video by looping through all the slides and set its play mode as auto.
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IVideo)
        {
            (shape as IVideo).PlayMode = VideoPlayMode.Auto;
        }
    }
}
```

---

# Spire.Presentation Speaker Notes Management
## Add and get speaker notes in PowerPoint slides using Spire.Presentation
```csharp
//Get the first slide from the presentation
ISlide slide = presentation.Slides[0];

//Get the NotesSlide in the first slide, if there is no notes, add it
NotesSlide ns = slide.NotesSlide;
if (ns == null)
{
    ns = slide.AddNotesSlide();
}

//Add text as the notes
ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation";

//Get the speaker notes
StringBuilder content = new StringBuilder();
content.AppendLine("The speaker notes added by Spire.Presentation is: " + ns.NotesTextFrame.Text);
```

---

# Spire.Presentation C# Comment
## Add comment to PowerPoint slide
```csharp
// Comment author
ICommentAuthor author = presentation.CommentAuthors.AddAuthor("E-iceblue", "comment:");

// Add comment
presentation.Slides[0].AddComment(author, "Add comment", new System.Drawing.PointF(18, 25), DateTime.Now);
```

---

# Spire.Presentation CSharp Add Notes
## Add notes to PowerPoint presentation slides
```csharp
// Get the first slide
ISlide slide = ppt.Slides[0];

// Add note slide
NotesSlide notesSlide = slide.AddNotesSlide();

// Add paragraph in the notesSlide
TextParagraph paragraph = new TextParagraph();
paragraph.Text = "Tips for making effective presentations:";
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);

paragraph = new TextParagraph();
paragraph.Text = "Use the slide master feature to create a consistent and simple design template.";
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);
// Set the bullet type for the paragraph in notesSlide
notesSlide.NotesTextFrame.Paragraphs[1].BulletType = TextBulletType.Numbered;
notesSlide.NotesTextFrame.Paragraphs[1].BulletStyle = NumberedBulletStyle.BulletArabicPeriod;

paragraph = new TextParagraph();
paragraph.Text = "Simplify and limit the number of words on each screen.";
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);
notesSlide.NotesTextFrame.Paragraphs[2].BulletType = TextBulletType.Numbered;
notesSlide.NotesTextFrame.Paragraphs[2].BulletStyle = NumberedBulletStyle.BulletArabicPeriod;

paragraph = new TextParagraph();
paragraph.Text = "Use contrasting colors for text and background.";
notesSlide.NotesTextFrame.Paragraphs.Append(paragraph);
notesSlide.NotesTextFrame.Paragraphs[3].BulletType = TextBulletType.Numbered;
notesSlide.NotesTextFrame.Paragraphs[3].BulletStyle = NumberedBulletStyle.BulletArabicPeriod;
```

---

# spire.presentation csharp comment manipulation
## delete and replace comments in PowerPoint presentation
```csharp
//Replace the text in the comment
presentation.Slides[0].Comments[1].Text = "Replace comment";

//Delete the third comment
presentation.Slides[0].DeleteComment(presentation.Slides[0].Comments[2]);
```

---

# spire.presentation csharp comment extraction
## extract comments from PowerPoint slides
```csharp
//Get all comments from the first slide.
Comment[] comments = presentation.Slides[0].Comments;

//Extract comment text
for (int i = 0; i < comments.Length; i++)
{
    str.Append(comments[i].Text + "\r\n");
}
```

---

# spire.presentation csharp slide comments
## retrieve comments from presentation slides
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load document from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Comments.pptx");

//Loop through comments
foreach (ICommentAuthor commentAuthor in presentation.CommentAuthors)
{
    foreach (Comment comment in commentAuthor.CommentsList)
    {
        //Get comment information
        string commentText = comment.Text;
        string authorName = comment.AuthorName;
        DateTime time = comment.DateTime;
        MessageBox.Show("Comment text : "+ comment.Text +"\n"+"Comment author : " + comment.AuthorName + "\n" + "Posted on time : " + comment.DateTime);
    }
}
```

---

# spire.presentation csharp powerpoint conversion
## convert PowerPoint to SVG while retaining notes
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_5.pptx");

//Retain the notes while converting PowerPoint file to svg file.
presentation.IsNoteRetained = true;

//Convert presentation slides to svg file.
Queue<byte[]> bytes = presentation.SaveToSVG();
```

---

# spire.presentation csharp slide notes
## remove notes from specific slide
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];

//Get note slide
NotesSlide note = slide.NotesSlide;
//Clear note text
note.NotesTextFrame.Text = "";
```

---

# Spire.Presentation Speaker Notes Removal
## Remove speaker notes from PowerPoint slides
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Get the first slide from the sample document.
ISlide slide = presentation.Slides[0];

//Remove the first speak note.
slide.NotesSlide.NotesTextFrame.Paragraphs.RemoveAt(1);
```

---

# Spire.Presentation C# Comment Reply
## Add and manage replies to comments in PowerPoint presentations
```csharp
//Create ppt file
Presentation ppt = new Presentation();        

//Create Comment author
ICommentAuthor author = ppt.CommentAuthors.AddAuthor("E-iceblue", "comment");

//Add comment
ppt.Slides[0].AddComment(author, "Add comment", new System.Drawing.Point(18, 25), DateTime.Now);
Comment comment = ppt.Slides[0].Comments[0];

//Add reply to Comment
if (!comment.IsReply)
{
    comment.Reply(author, "Add Reply1", DateTime.Now);
    comment.Reply(author, "Add Reply2", DateTime.Now);
}

//delete first reply
ppt.Slides[0].DeleteComment(author, "Add Reply1");
```

---

# spire.presentation csharp header footer
## set header and footer in powerpoint presentation
```csharp
//Add footer
presentation.SetFooterText("Demo of Spire.Presentation");

//Set the footer visible
presentation.FooterVisible = true;

//Set the page number visible
presentation.SlideNumberVisible = true;

//Set the date visible
presentation.DateTimeVisible = true;
```

---

# Spire.Presentation C# Note Master Header Footer
## Manage header and footer in note master slides
```csharp
//Set the note Masters header and footer
INoteMasterSlide noteMasterSlide = presentation.NotesMaster;
if (!noteMasterSlide.Equals(null))
{
    foreach(Shape shape in noteMasterSlide.Shapes)
    {
        if (!shape.Placeholder.Equals(null))
        {
            if (shape.Placeholder.Type.Equals(PlaceholderType.Header))
            {
                (shape as IAutoShape).TextFrame.Text = "change the header by Spire";
            }
            if (shape.Placeholder.Type.Equals(PlaceholderType.Footer))
            {
                (shape as IAutoShape).TextFrame.Text = "change the footer by Spire";
            }
        }
    }
}
```

---

# spire.presentation csharp smartart
## access child nodes of smartart in powerpoint
```csharp
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        //Get the SmartArt and collect nodes
        ISmartArt sa = shape as ISmartArt;
        ISmartArtNodeCollection nodes = sa.Nodes;
        
        //Access the parent node at position 0
        ISmartArtNode node = nodes[0];
        //Traverse through all child nodes inside SmartArt
        for (int i = 0; i < node.ChildNodes.Count; i++)
        {
            //Access SmartArt child node at index i
            ISmartArtNode childnode = node.ChildNodes[i];
            //Get the SmartArt child node parameters
            string nodeText = childnode.TextFrame.Text;
            int nodeLevel = childnode.Level;
            int nodePosition = childnode.Position;
        }
    }
}
```

---

# spire.presentation csharp smartart
## access SmartArt nodes in PowerPoint presentation
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT
presentation.LoadFromFile("SmartArt.pptx");

StringBuilder strB = new StringBuilder();
strB.AppendLine("Access SmartArt nodes.");
strB.AppendLine("Here is the SmartArt node parameters details:"); 
string outString="";
ISmartArtNode node;

foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        //Get the SmartArt
        ISmartArt sa = shape as ISmartArt;
        
        ISmartArtNodeCollection nodes = sa.Nodes;
    
        //Traverse through all nodes inside SmartArt
        for (int i = 0; i < nodes.Count; i++)
        {
            //Access SmartArt node at index i
            node = nodes[i];
            //Print the SmartArt node parameters
            outString = string.Format("Node text = {0}, Node level = {1}, Node Position = {2}", node.TextFrame.Text, node.Level, node.Position);
            strB.AppendLine(outString);
        }
    }
}
```

---

# Spire.Presentation C# SmartArt
## Access SmartArt layout type from PowerPoint presentation
```csharp
// Iterate through shapes in the first slide
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        // Get the SmartArt
        ISmartArt sa = shape as ISmartArt;
        // Check SmartArt Layout
        String layout = sa.LayoutType.ToString();
        MessageBox.Show("SmartArt layout type is " + layout);
    }
}
```

---

# Spire.Presentation C# SmartArt Node Access
## Access specific child node in SmartArt and get its properties
```csharp
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        //Get the SmartArt
        ISmartArt sa = shape as ISmartArt;

        //Get SmartArt node collection 
        ISmartArtNodeCollection nodes = sa.Nodes;

        //Access SmartArt node at index 0
        ISmartArtNode node = nodes[0];

        //Access SmartArt child node at index 1
        ISmartArtNode childNode = node.ChildNodes[1];

        //Print the SmartArt child node parameters
        string outString = string.Format("Node text = {0}, Node level = {1}, Node Position = {2}", childNode.TextFrame.Text, childNode.Level, childNode.Position);
    }
}
```

---

# Spire.Presentation C# SmartArt
## Add SmartArt nodes by position in PowerPoint presentation
```csharp
// Check if shape is SmartArt
if (shape is ISmartArt)
{
    // Get the SmartArt
    ISmartArt smartArt = shape as ISmartArt;

    int position = 0;
    // Add a new node at specific position
    ISmartArtNode node = smartArt.Nodes.AddNodeByPosition(position);
    // Add text and set the text style 
    node.TextFrame.Text = "New Node";
    node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Red;

    // Get a node
    node = smartArt.Nodes[1];                
    position = 1;
    // Add a new child node at specific position
    ISmartArtNode childNode = node.ChildNodes.AddNodeByPosition(position);
    // Add text and set the text style 
    childNode.TextFrame.Text = "New child node";
    childNode.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
    childNode.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Blue;
}
```

---

# spire.presentation csharp smartart
## Add SmartArt node to PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile("AddSmartArtNode.pptx");

//Get the SmartArt
ISmartArt sa = presentation.Slides[0].Shapes[0] as ISmartArt;

//Add a node
ISmartArtNode node = sa.Nodes.AddNode();
//Add text and set the text style 
node.TextFrame.Text = "AddText";
node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink;

presentation.SaveToFile("AddSmartArtNode.pptx", FileFormat.Pptx2010);
```

---

# Spire.Presentation SmartArt Assistant Node
## Set SmartArt nodes as assistant nodes in PowerPoint presentation
```csharp
//Get the SmartArt and collect nodes
ISmartArt smartArt = shape as ISmartArt;

ISmartArtNodeCollection nodes = smartArt.Nodes;

//Traverse through all nodes inside SmartArt
for (int i = 0; i < nodes.Count; i++)
{
    //Access SmartArt node at index i
    node = nodes[i];
    // Check if node is assitant node
    if (!node.IsAssistant)
    {
        //Set node as assitant node
        node.IsAssistant = true;
    }
}
```

---

# spire.presentation csharp smartart
## change SmartArt node text
```csharp
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        //Get the SmartArt and collect nodes
        ISmartArt smartArt = shape as ISmartArt;
        //Obtain the reference of a node by using its Index  
        // select second root node
        ISmartArtNode node = smartArt.Nodes[1]; 
        // Set the text of the TextFrame 
        node.TextFrame.Text = "Second root node";
    }
}
```

---

# Spire.Presentation C# SmartArt
## Change SmartArt color style
```csharp
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        //Get the SmartArt
        ISmartArt smartArt = shape as ISmartArt;
        // Check SmartArt color type
        if (smartArt.ColorStyle == SmartArtColorType.ColoredFillAccent1)
        {
            // Change SmartArt color type
            smartArt.ColorStyle = SmartArtColorType.ColorfulAccentColors;
        }
    }
}
```

---

# spire.presentation csharp smartart
## change SmartArt shape style in PowerPoint presentation
```csharp
// Iterate through shapes on the first slide
foreach (IShape shape in presentation.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        // Get the SmartArt
        ISmartArt smartArt = shape as ISmartArt;
        
        // Check SmartArt style
        if (smartArt.Style == SmartArtStyleType.SimpleFill)
        {
            // Change SmartArt Style
            smartArt.Style = SmartArtStyleType.Cartoon;
        }
    }
}
```

---

# Spire Presentation SmartArt Creation
## Create and configure SmartArt shapes in PowerPoint presentations
```csharp
//Create a SmartArt shape on the first slide
Spire.Presentation.Diagrams.ISmartArt sa = presentation.Slides[0].Shapes.AppendSmartArt(200, 60, 300, 300, Spire.Presentation.Diagrams.SmartArtLayoutType.Gear);

//Set type and color of smartart
sa.Style = Spire.Presentation.Diagrams.SmartArtStyleType.SubtleEffect;
sa.ColorStyle = Spire.Presentation.Diagrams.SmartArtColorType.GradientLoopAccent3;

//Remove all shapes
foreach (object a in sa.Nodes)
    sa.Nodes.RemoveNode(0);

//Add two custom shapes with text
Spire.Presentation.Diagrams.ISmartArtNode node = sa.Nodes.AddNode();
sa.Nodes[0].TextFrame.Text = "aa";
node = sa.Nodes.AddNode();
node.TextFrame.Text = "bb";
node.TextFrame.TextRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Black;
```

---

# Extract Text from SmartArt in PowerPoint
## This code demonstrates how to extract text from SmartArt shapes in a PowerPoint presentation.
```csharp
//Traverse through all the slides of the PPT file and find the SmartArt shapes.
StringBuilder st = new StringBuilder();
for (int i = 0; i < presentation.Slides.Count; i++)
{
    for (int j = 0; j < presentation.Slides[i].Shapes.Count; j++)
    {
        if (presentation.Slides[i].Shapes[j] is ISmartArt)
        {
            ISmartArt smartArt = presentation.Slides[i].Shapes[j] as ISmartArt;

            //Extract text from SmartArt and append to the StringBuilder object.
            for (int k = 0; k < smartArt.Nodes.Count; k++)
            {
                st.AppendLine(smartArt.Nodes[k].TextFrame.Text);
            }
        }
    }
}
```

---

# spire.presentation csharp smartart
## remove a node from smartart in powerpoint
```csharp
//Get the SmartArt and collect nodes
ISmartArt sa = presentation.Slides[0].Shapes[0] as ISmartArt;
ISmartArtNodeCollection nodes = sa.Nodes;

//Remove the node to specific position
nodes.RemoveNodeByPosition(2);
```

---

# Spire.Presentation SmartArt Link Line Outline
## Set outline properties for SmartArt link lines in PowerPoint
```csharp
//Get the specified shape as ISmartArt
ISmartArt smartArt = ppt.Slides[0].Shapes[0] as ISmartArt;
int count = smartArt.Nodes.Count;
ISmartArtNode node;
//Loop through all smartArts
for (int i = 0; i < count; i++)
{
    node = smartArt.Nodes[i];
    //Set the line type
    node.LinkLine.FillType = FillFormatType.Solid;
    //Set the line color
    node.LinkLine.SolidFillColor.Color = Color.Red;
    //Set the line width
    node.LinkLine.Width = 2;
    //Set the line DashStyle
    node.LinkLine.DashStyle = LineDashStyleType.SystemDash;
}
```

---

# Spire.Presentation C# SmartArt
## Set SmartArt node outline properties
```csharp
//Set ISmartArt form special shape
ISmartArt smartArt = ppt.Slides[0].Shapes[0] as ISmartArt;
int count = smartArt.Nodes.Count;
ISmartArtNode node;
//Loop through all nodes
for (int i = 0; i < count; i++)
{
    node = smartArt.Nodes[i];
    //Set the fill format type
    node.Line.FillType = FillFormatType.Solid;
    //Set the line style
    node.Line.Style = TextLineStyle.ThinThin;
    //Set the line color
    node.Line.SolidFillColor.Color = Color.Red;
    //Set the line width
    node.Line.Width = 2;
}
```

---

# Spire.Presentation C# Watermark
## Add image watermark to PowerPoint slide
```csharp
//Set the properties of SlideBackground, and then fill the image as watermark.
presentation.Slides[0].SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom;
presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Picture;
presentation.Slides[0].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch;
presentation.Slides[0].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image;
```

---

# Spire.Presentation C# Watermark
## Add watermark to PowerPoint slides
```csharp
//Get the size of the watermark string
Graphics gc = this.CreateGraphics();
SizeF size = gc.MeasureString("E-iceblue", new Font("Lucida Sans Unicode", 50));

//Define a rectangle range
RectangleF rect = new RectangleF((presentation.SlideSize.Size.Width - size.Width) / 2, (presentation.SlideSize.Size.Height - size.Height) / 2, size.Width, size.Height);

//Add a rectangle shape with a defined range
IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, rect);

//Set the style of the shape
shape.Fill.FillType = FillFormatType.None;
shape.ShapeStyle.LineColor.Color = Color.White;
shape.Rotation = -45;
shape.Locking.SelectionProtection = true;
shape.Line.FillType = FillFormatType.None;

//Add text to the shape
shape.TextFrame.Text = "E-iceblue";
TextRange textRange = shape.TextFrame.TextRange;
//Set the style of the text range
textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
textRange.Fill.SolidColor.Color = Color.FromArgb(120, Color.HotPink);
textRange.FontHeight = 50;
```

---

# spire.presentation csharp watermark removal
## Remove text and image watermarks from PowerPoint presentations
```csharp
//Remove text watermark by removing the shape which contains the text string "E-iceblue".
for (int i = 0; i < presentation.Slides.Count; i++)
{
    for (int j = 0; j < presentation.Slides[i].Shapes.Count; j++)
    {
        if (presentation.Slides[i].Shapes[j] is IAutoShape)
        {
            IAutoShape shape = presentation.Slides[i].Shapes[j] as IAutoShape;
            if (shape.TextFrame.Text.Contains("E-iceblue"))
            {
                presentation.Slides[i].Shapes.Remove(shape);
            }
        }
    }
}

//Remove image watermark.
for (int i = 0; i < presentation.Slides.Count; i++)
{
    presentation.Slides[i].SlideBackground.Fill.FillType = FillFormatType.None;
}
```

---

# Spire.Presentation OLE Embedding
## Embed Excel file as OLE object in PowerPoint presentation
```csharp
//Create a Presentaion document
Presentation ppt = new Presentation();

//Load a image file
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\EmbedExcelAsOLE.png");
IImageData oleImage = ppt.Images.Append(image);
Rectangle rec = new Rectangle(80, 60, image.Width, image.Height);

//Insert an OLE object to presentation based on the Excel data
Spire.Presentation.IOleObject oleObject = ppt.Slides[0].Shapes.AppendOleObject("excel", File.ReadAllBytes(@"..\..\..\..\..\..\Data\EmbedExcelAsOLE.xlsx"), rec);
oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
oleObject.ProgId = "Excel.Sheet.12";
```

---

# Embed ZIP into PowerPoint
## Core functionality to embed a ZIP file as an OLE object in a PowerPoint presentation
```csharp
// Create a presentation
Presentation ppt = new Presentation();

Rectangle rec = new Rectangle(80, 60, 100, 100);

// Insert the zip object to presentation
// zipData should contain the byte array of the ZIP file
IOleObject ole = ppt.Slides[0].Shapes.AppendOleObject(@"test.zip", zipData, rec);
ole.ProgId = "Package";
// iconImage should contain the image to be displayed as the OLE object icon
IImageData oleImage = ppt.Images.Append(iconImage);
ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
```

---

# spire.presentation csharp ole extraction
## extract OLE objects from PowerPoint presentation
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load document from disk
presentation.LoadFromFile("ExtractOLEObject.pptx");

//Loop through the slides and shapes
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IOleObject)
        {
            //Find OLE object
            IOleObject oleObject = shape as IOleObject;

            //Get its data and write to file
            byte[] bytes = oleObject.Data;
            switch (oleObject.ProgId)
            {
                case "Excel.Sheet.8":
                    File.WriteAllBytes("result.xls", bytes);
                    break;
                case "Excel.Sheet.12":
                    File.WriteAllBytes("result.xlsx", bytes);
                    break;
                case "Word.Document.8":
                    File.WriteAllBytes("result.doc", bytes);
                    break;
                case "Word.Document.12":
                    File.WriteAllBytes("result.docx", bytes);
                    break;
                case "PowerPoint.Show.8":
                    File.WriteAllBytes("result.ppt", bytes);
                    break;
                case "PowerPoint.Show.12":
                    File.WriteAllBytes("result.pptx", bytes);
                    break;
            }
        }
    }
}
```

---

# spire.presentation csharp ole properties
## extract OLE object properties from PowerPoint slides
```csharp
//Get the first slide
ISlide slide = presentation.Slides[0];

//Get the first OLE
OleObjectCollection oles = slide.OleObjects;
OleObject oleObject = oles[0];

//Get the information of OLE Object
oleObject.ShapeID;
oleObject.Frame.Top;
oleObject.Frame.Left;
oleObject.Frame.Width;
oleObject.Frame.Height;
oleObject.IsHidden;

//Get the properties of OLE
foreach (DictionaryEntry entry in oleObject.Properties)
{
    // Access OLE property key and value
    entry.Key;
    entry.Value;
}
```

---

# spire.presentation csharp ole modification
## modify OLE object data in PowerPoint presentation
```csharp
//Loop through the slides and shapes
foreach (ISlide slide in presentation.Slides)
{
    foreach (IShape shape in slide.Shapes)
    {
        if (shape is IOleObject)
        {
            //Find OLE object
            IOleObject oleObject = shape as IOleObject;

            //Get its data and write to file
            byte[] bytes = oleObject.Data;
            MemoryStream pptStream = new MemoryStream(bytes);
            MemoryStream stream=new MemoryStream();
            if (oleObject.ProgId == "PowerPoint.Show.12")
            {
                //Load the PPT stream
                Presentation ppt = new Presentation();
                ppt.LoadFromStream(pptStream, Spire.Presentation.FileFormat.Auto);
                //Append an image in slide
                ppt.Slides[0].Shapes.AppendEmbedImage(ShapeType.Rectangle, @"..\..\..\..\..\..\Data\Logo.png", new RectangleF(50, 50, 100, 100));
                ppt.SaveToFile(stream, Spire.Presentation.FileFormat.Pptx2013);
                stream.Position = 0;
                //Modify the data
                oleObject.Data = stream.ToArray();
            }
        }
    }
}
```

---

# spire.presentation csharp print
## print presentation document with custom printer settings
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.LoadFromFile("path_to_presentation_file");

//Print
PrinterSettings printerSettings = new PrinterSettings();
printerSettings.FromPage = 0;
printerSettings.ToPage = presentation.Slides.Count-1;
presentation.Print(printerSettings);
```

---

# spire.presentation csharp print
## print multiple slides into one page
```csharp
//Create a PPT document
Presentation ppt = new Presentation();
//Load the document from disk
ppt.LoadFromFile("PrintMultipleSlidesIntoOnePage.pptx");
PresentationPrintDocument document = new PresentationPrintDocument(ppt);

//Set print task name
document.DocumentName = "print task 1";
document.PrintOrder = Order.Horizontal;
document.SlideFrameForPrint = true;

//Set Gray level when printing
document.GrayLevelForPrint = true;
//Set four slides on one page
document.SlideCountPerPageForPrint = PageSlideCount.Four;

//Set continuous print area
document.PrinterSettings.PrintRange = PrintRange.SomePages;
document.PrinterSettings.FromPage = 1;
document.PrinterSettings.ToPage = ppt.Slides.Count - 1;

//Set discontinuous print area
//document.SelectSldiesForPrint("1", "2-4");

ppt.Print(document);
ppt.Dispose();
```

---

# Spire Presentation C# Print
## Print PowerPoint presentation to virtual printer
```csharp
// Create a PowerPoint document
Presentation presentation = new Presentation();

// Load the file from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Ppt_6.pptx");

// Print PowerPoint document to virtual printer (Microsoft XPS Document Writer)
PresentationPrintDocument document = new PresentationPrintDocument(presentation);
document.PrinterSettings.PrinterName = "Microsoft XPS Document Writer";

presentation.Print(document);
```

---

# spire.presentation csharp print
## print specified range of PowerPoint pages
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Create a print document for the presentation
PresentationPrintDocument document = new PresentationPrintDocument(presentation);

//Set the document name to display while printing the document
document.DocumentName = "Template_Ppt_6.pptx";

//Choose to print some pages from the PowerPoint document
document.PrinterSettings.PrintRange = PrintRange.SomePages;
document.PrinterSettings.FromPage = 2;
document.PrinterSettings.ToPage = 3;

//Set the number of copies of the document to print
document.PrinterSettings.Copies = 2;

//Print the document
presentation.Print(document);
```

---

# spire.presentation print settings
## configure print settings for PowerPoint presentation
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Use PrintDocument object to print presentation slides.
PresentationPrintDocument document = new PresentationPrintDocument(presentation);

//Print document to virtual printer.
document.PrinterSettings.PrinterName = "Microsoft XPS Document Writer";

//Print the slide with frame.
presentation.SlideFrameForPrint = true;

//Print 4 slides horizontal.
presentation.SlideCountPerPageForPrint = PageSlideCount.Four;
presentation.OrderForPrint = Order.Horizontal;

//Print the slide with Grayscale.
presentation.GrayLevelForPrint = true;

//Set the print document name.          
document.DocumentName = "Template_Ppt_6.pptx";

document.PrinterSettings.PrintToFile = true;
document.PrinterSettings.PrintFileName = ("Result-SetPrintSettingsByPrintDocumentObject.xps");

presentation.Print(document);
```

---

# Spire.Presentation C# Print Settings
## Set print settings for a presentation using PrinterSettings object
```csharp
//Use PrinterSettings object to print presentation slides.
PrinterSettings ps = new PrinterSettings();
ps.PrintRange = PrintRange.AllPages;
ps.PrintToFile = true;
String result = "Result-SetPrintSettingsByPrinterSettingsObject.xps";
ps.PrintFileName = (result);

//Print the slide with frame.
presentation.SlideFrameForPrint = true;

//Print the slide with Grayscale.
presentation.GrayLevelForPrint = true;

//Print 4 slides horizontal.
presentation.SlideCountPerPageForPrint = PageSlideCount.Four;
presentation.OrderForPrint = Order.Horizontal;

//Only select some slides to print.
//presentation.SelectSlidesForPrint("1", "3");

//Print the document.
presentation.Print(ps);
```

---

# Spire.Presentation C# Print
## Silently print PowerPoint presentation using default printer
```csharp
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.LoadFromFile("file_path.pptx");

//Print the PowerPoint document to default printer.
PresentationPrintDocument document = new PresentationPrintDocument(presentation);
document.PrintController = new StandardPrintController();

presentation.Print(document);
```

---

# spire.presentation csharp print
## print presentation to specific printer
```csharp
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk.
presentation.LoadFromFile("presentation.pptx");

//New PrintSeetings
PrinterSettings printerSettings = new PrinterSettings();

//Set landscape for page
printerSettings.DefaultPageSettings.Landscape = true;

//Specific the printer
printerSettings.PrinterName = "Microsoft XPS Document Writer";

//Print 
presentation.Print(printerSettings);
```

---

# Spire.Presentation C# VBA Macro Removal
## Remove VBA macros from PowerPoint presentations using Spire.Presentation library
```csharp
//Create a PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.LoadFromFile(@"..\..\..\..\..\..\Data\Macros.ppt");
//Remove macros
//Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
presentation.DeleteMacros();
string result = "RemoveVBAMacros_result.ppt";
presentation.SaveToFile(result,FileFormat.PPT);
```

---

