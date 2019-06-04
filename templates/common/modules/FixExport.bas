Attribute VB_Name = "FixExport"
' See https://github.com/htwchur/scroll-exporter-templates/issues/14
Sub FixTableOfContents()
    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next
End Sub
Sub SetDocumentPropertiesFromShapeContents()
    Call SetDocumentPropertyFromShape(wdPropertyTitle, "title")
    Call SetDocumentPropertyFromShape(wdPropertySubject, "title")
    Call SetDocumentPropertyFromShape(wdPropertyAuthor, "author")
    Call SetDocumentPropertyFromShape(wdPropertyCategory, "classification")
    Call SetDocumentPropertyFromShape(wdPropertyCompany, "scope")
    Call SetDocumentPropertyFromShape(wdPropertyManager, "issuingOffice")
End Sub
Function GetShape(name As String) As Shape
    Dim oShape As Shape
    For Each oShape In ActiveDocument.Shapes
        If (oShape.name = name) Then
            Set GetShape = oShape
            Exit Function
        End If
    Next
    Set GetShape = Nothing
End Function
Sub SetDocumentPropertyFromShape(id, name As String)
    Dim oShape As Shape
    Dim text As String
    
    Set oShape = GetShape(name)
    If Not oShape Is Nothing Then
        text = oShape.TextFrame.TextRange.text
        ActiveDocument.BuiltInDocumentProperties(id) = text
    End If
End Sub


Sub FixPlaceholders(section As Integer, style As String)
    Dim shp As Shape
    Dim str As String

    For Each shp In ActiveDocument.Sections(section).Range.ShapeRange
        ' only fix text boxes
        If (shp.Type = MsoShapeType.msoTextBox) Then
            
            shp.Select
            str = Selection.ShapeRange.TextFrame.TextRange.text
            
            ' Only first line
            str = Split(str, Chr(13))(0)
            ' Debug.Print (str)
    
            shp.TextFrame.TextRange.Delete
            shp.TextFrame.DeleteText
            shp.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            shp.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
            shp.TextFrame.TextRange.text = str
            
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