Attribute VB_Name = "ScrollOffice"
Public Function StandardDictionary() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    Call dic.Add("title", "$scroll.title")
    Set StandardDictionary = dic
End Function

Public Function PagePropertiesDictionary() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    Call dic.Add("title", "$scroll.title")
    Call dic.Add("author", "$scroll.pageproperty.(Autor)")
    Call dic.Add("issuingOffice", "$scroll.pageproperty.(Ausgabestelle)")
    Call dic.Add("scope", "$scroll.pageproperty.(Geltungsbereich)")
    Call dic.Add("classification", "$scroll.pageproperty.(Klassifizierung)")
    Call dic.Add("version", "$scroll.pageproperty.(Version)")
    Call dic.Add("issuingDate", "$scroll.pageproperty.(Ausgabedatum)")
    Call dic.Add("distribution", "$scroll.pageproperty.(Verteiler)")
    Set PagePropertiesDictionary = dic
End Function

Public Function ConfluenceDictionary() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    Call dic.Add("title", "$scroll.title")
    Call dic.Add("author", "$scroll.modifier.fullName")
    Call dic.Add("issuingOffice", "$scroll.space.name")
    Call dic.Add("scope", "$scroll.space.name")
    Call dic.Add("classification", "Intern")
    Call dic.Add("version", "$scroll.version")
    Call dic.Add("issuingDate", "$scroll.modificationdate")
    Call dic.Add("distribution", "-")
    Set ConfluenceDictionary = dic
End Function


Public Sub Replace(ByRef cc As ContentControl, ByRef dic As Dictionary)
    Dim tV As Variant
    tV = cc.tag
    
    Dim t As String
    t = CStr(tV)
    Debug.Print (t)
    
    Dim v As String
    If dic.Exists(t) Then
       
       v = dic.Item(t)
    
       Dim r As Range
       Set r = cc.Range
           
       cc.Delete
       r.Delete
       r.InsertAfter (v)
    End If
End Sub
    
    
Sub ReplaceContentControls(ByRef doc As Document, ByRef dic As Dictionary)
    ' https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm
    Dim rngStory As Word.Range
    Dim lngJunk As Long
    'Fix the skipped blank Header/Footer problem as provided by Peter Hewett
    lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
    'Iterate through all story types in the current document
    For Each rngStory In ActiveDocument.StoryRanges
        'Iterate through all linked stories
        Do
            Dim cc As ContentControl
            For Each cc In rngStory.ContentControls
                Call Replace(cc, dic)
            Next
            'Get next linked story (if any)
            Set rngStory = rngStory.NextStoryRange
        Loop Until rngStory Is Nothing
    Next
End Sub
    
Sub Replace3EndWithScrollContent(ByRef doc As Document)
    '
    Call Selection.GoTo(wdGoToPage, wdGoToAbsolute, 3)
    Call Selection.EndKey(wdStory, wdExtend)
    Dim r As Range
    Set r = Selection.Range
    r.Delete
    r.InsertAfter ("$scroll.content")

End Sub

Sub ConvertToPageProperties(doc As Word.Document)
    Dim dic As Dictionary
    Set dic = PagePropertiesDictionary()
    
    Call ReplaceContentControls(doc, dic)
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls(doc, dic)
    Call ReplaceContentControls(doc, dic)
    
    Call Replace3EndWithScrollContent(doc)
    
    Call CreateScrollOfficeStyles(doc)
  
    Call ReplaceTables(doc)
    
End Sub

Sub ConvertToStandard(ByRef doc As Document)
    Dim dic As Dictionary
    Set dic = StandardDictionary()
    
    Call ReplaceContentControls(doc, dic)
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls(doc, dic)
    Call ReplaceContentControls(doc, dic)

    Call Replace3EndWithScrollContent(doc)
    
    Call CreateScrollOfficeStyles(doc)
End Sub

Sub ConvertToConfluence(ByRef doc As Document)
Dim dic As Dictionary
    Set dic = ConfluenceDictionary()
    
    Call ReplaceContentControls(doc, dic)
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls(doc, dic)
    Call ReplaceContentControls(doc, dic)

    Call Replace3EndWithScrollContent(doc)
    
    Call CreateScrollOfficeStyles(doc)

End Sub

Sub CreateScrollOfficeStyles(ByRef doc As Document)
    
    Call CreateOrEditStyle(doc, "Scroll List Bullet", "Aufzählungszeichen")
    Call CreateOrEditStyle(doc, "Scroll List Bullet 1", "Aufzählungszeichen")
    Call CreateOrEditStyle(doc, "Scroll List Bullet 2", "Aufzählungszeichen 2")
    Call CreateOrEditStyle(doc, "Scroll List Bullet 3", "Aufzählungszeichen 3")
    Call CreateOrEditStyle(doc, "Scroll List Bullet 4", "Aufzählungszeichen 4")
    Call CreateOrEditStyle(doc, "Scroll List Bullet 5", "Aufzählungszeichen 5")
    
    Call CreateOrEditStyle(doc, "Scroll List Number", "Listennummer")
    Call CreateOrEditStyle(doc, "Scroll List Number 1", "Listennummer")
    Call CreateOrEditStyle(doc, "Scroll List Number 2", "Listennummer 2")
    Call CreateOrEditStyle(doc, "Scroll List Number 3", "Listennummer 3")
    Call CreateOrEditStyle(doc, "Scroll List Number 4", "Listennummer 4")
    Call CreateOrEditStyle(doc, "Scroll List Number 5", "Listennummer 5")
    
    Call CreateOrEditTableStyle(doc, "Scroll Table Normal", "Tabelle HTW Chur")

    Dim oStyle As Style
    petrolLight = RGB(217, 233, 237)
    ockerLight = RGB(240, 232, 227)
    redLight = RGB(244, 226, 226)
    
    Call CreateOrEditTable(doc, "Scroll Tip", petrolLight)
    Call CreateOrEditTable(doc, "Scroll Info", ockerLight)
    Call CreateOrEditTable(doc, "Scroll Note", ockerLight)
    Call CreateOrEditTable(doc, "Scroll Warning", redLight)
    
End Sub

Sub CreateOrEditTable(ByRef doc As Document, styleName As String, color)
    Set oStyle = CreateOrEditTableStyle(doc, styleName, "Normale Tabelle")
    oStyle.Table.Shading.BackgroundPatternColor = color
End Sub


Function CreateOrEditTableStyle(ByRef doc As Document, styleName As String, baseStyleName As String) As Style
    Dim oStyle As Style
    If StyleExists(styleName) Then
        Debug.Print (doc.Styles(styleName))
        Set oStyle = doc.Styles(styleName)
    Else
        Set oStyle = doc.Styles.Add(name:=styleName, Type:=WdStyleType.wdStyleTypeTable)
    End If
    oStyle.BaseStyle = baseStyleName
    Set CreateOrEditTableStyle = oStyle
End Function


Function CreateOrEditStyle(ByRef doc As Document, styleName As String, baseStyleName As String) As Style
    Dim oStyle As Style
    If StyleExists(styleName) Then
        Set oStyle = doc.Styles(styleName)
    Else
        Set oStyle = doc.Styles.Add(styleName, WdStyleType.wdStyleTypeParagraphOnly)
    End If
    oStyle.BaseStyle = baseStyleName
    Set CreateOrEditStyle = oStyle
End Function

Function StyleExists(styleName As String) As Boolean
    Dim oStyle As Style
    StyleExists = False
    For Each oStyle In ActiveDocument.Styles
        If oStyle.NameLocal = styleName Then
            StyleExists = True
            Exit Function
        End If
    Next oStyle
    Exit Function
End Function


Function ReplaceTables(ByRef doc As Document)
    Dim r As Range
    
    doc.Bookmarks("ChangeControl_Range").Range.Delete
    doc.Bookmarks("ChangeControl_Start").Range.InsertAfter ("$scroll.pageproperty.(Aenderungskontrolle)" + vbNewLine + vbNewLine + "$scroll.pageproperty.(Freigabe)")
    
End Function
