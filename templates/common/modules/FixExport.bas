Attribute VB_Name = "FixExport"
' See https://github.com/htwchur/scroll-exporter-templates/issues/14
Sub FixTableOfContents()
    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next
End Sub
Sub FixPlaceholders(section As Integer, style As String)
    Dim shp As Shape
    Dim str As String

    For Each shp In ActiveDocument.Sections(section).Range.ShapeRange
        ' only fix text boxes
        If (shp.Type = MsoShapeType.msoTextBox) Then

            shp.Select
            str = Selection.ShapeRange.TextFrame.TextRange.Text
            
            ' Only first line
            str = Split(str, Chr(13))(0)
            ' Debug.Print (str)
    
            shp.TextFrame.TextRange.Delete
            shp.TextFrame.DeleteText
            shp.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            shp.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
            shp.TextFrame.TextRange.Text = str
            
            ' set formating
            shp.TextFrame.TextRange.Paragraphs(1).style = style
        End If
    Next
End Sub

Sub FixBold(section As Integer)
    Set myRange = ActiveDocument.Sections(section).Range
    Set oFind = myRange.Find
    oFind.ClearFormatting
    oFind.Font.Bold = True
    ' oFind.Style = ActiveDocument.Styles("Hervorhebung")
    oFind.Text = ""
    oFind.Forward = True
    oFind.Format = True
    With oFind.Replacement
        .ClearFormatting
        .Font.Bold = False
        .style = ActiveDocument.Styles("Intensive Hervorhebung")
    End With
    oFind.Execute FindText:="", ReplaceWith:="", Format:=True, Replace:=wdReplaceAll
End Sub

