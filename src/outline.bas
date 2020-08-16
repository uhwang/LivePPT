Sub outline()
Dim sld As Slide
Dim shp As Shape

For Each sld In ActivePresentation.Slides
    sld.Select
    ActiveWindow.ViewType = ppViewSlide
    ActiveWindow.Activate
    
    sld.Shapes(1).Select
    With ActiveWindow.Selection.TextRange2.Font
        'msoFalse doesn't work at all
        .Line.Visible = msoCTrue
        '.Line.Pattern = msoPattern10Percent
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = 0.5
        .Line.Style = msoLineSingle
        .Line.DashStyle = msoLineSolid
        .Line.Transparency = 0
    End With
Next sld
End Sub