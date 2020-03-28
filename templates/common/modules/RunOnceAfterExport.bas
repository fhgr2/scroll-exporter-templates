Attribute VB_Name = "RunOnceAfterExport"
Function ShouldRunOnceAfterExport() As Boolean
    ShouldRunOnceAfterExport = IsExported() And Not (hasRun())
End Function

Sub SetShouldRunOnceAfterExport(hasRun As Boolean)
    SetRun (hasRun)
End Sub

Public Function IsExported() As Boolean
    Dim oShape As Shape
    Set oShape = GetShape("title")
    If Not oShape Is Nothing And oShape.TextFrame.TextRange.Find.Execute("$scroll.title") Then
        IsExported = False
        Exit Function
    End If
    If ActiveDocument.Range.Find.Execute("$scroll.title") Then
        IsExported = False
        Exit Function
    End If
    IsExported = True
End Function

Private Sub EnsureHasBooleanCustomPropertyIfNeeded(name As String)
    If ActiveDocument.CustomDocumentProperties.Count = 0 Then
        With ActiveDocument.CustomDocumentProperties
        .Add name:=name, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=False
    End With
    End If

Dim dp As DocumentProperty

For Each dp In ActiveDocument.CustomDocumentProperties
    If (dp.name = name) Then
        Exit Sub
    End If
Next dp

With ActiveDocument.CustomDocumentProperties
    .Add name:=name, _
    LinkToContent:=False, _
    Type:=msoPropertyTypeBoolean, _
    Value:=False
End With

End Sub

Function GetCustomProperty(name As String) As DocumentProperty
    Call EnsureHasBooleanCustomPropertyIfNeeded(name)

    For Each dp In ActiveDocument.CustomDocumentProperties
        If (dp.name = name) Then
            Set GetCustomProperty = dp
            Exit Function
        End If
    Next dp

End Function


Private Function hasRun() As Boolean
    Dim name As String
    name = "HasRun"
    
    Dim dp As DocumentProperty
    Set dp = GetCustomProperty(name)

    hasRun = dp.Value
End Function

Public Sub SetRun(hasRun As Boolean)

    Dim name As String
    name = "HasRun"

Dim dp As DocumentProperty
Set dp = GetCustomProperty(name)

dp.Value = hasRun

End Sub

