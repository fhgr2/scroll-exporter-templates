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

