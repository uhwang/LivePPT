Attribute VB_Name = "Module1"
Sub shadow()
Dim sld As Slide
Dim shp As Shape
For Each sld In ActivePresentation.Slides
For Each shp In sld.Shapes
    With shp.TextFrame2.TextRange.Font
    .shadow.Visible = msoCTrue
    .shadow.Style = msoShadowStyleOuterShadow
    .shadow.OffsetX = 2
    .shadow.OffsetY = 2
    ' !!!! Do not use size option !!!!
    '.shadow.Size = 1
    .shadow.Blur = 2
    .shadow.Transparency = 0.7
    End With
Next shp
Next sld

End Sub
