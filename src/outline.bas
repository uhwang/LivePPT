Sub outline()
Dim sld As Slide
Dim shp As Shape

For Each sld In ActivePresentation.Slides
    sld.Select
    ActiveWindow.ViewType = ppViewSlide
    ActiveWindow.Activate
    
    sld.Shapes(1).Select
    With ActiveWindow.Selection.TextRange2.Font
        .Line.Visible = msoCTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = 2
    End With
Next sld
End Sub
