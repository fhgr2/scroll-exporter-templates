Attribute VB_Name = "Regulation_Macros"
Sub ExitReadingLayout()
    If IsExported() Then
        If ShouldRunOnceAfterExport() Then
            Call FixExport
            Call FixIdentification
            Call FixTables
            SetRun (True)
        Else
            SetRun (False)
        End If
    End If

End Sub

Sub FixExport()
    
    Const FirstLegalParagraphWithoutNumber As Boolean = False
    
    ' Selection.WholeStory
    ' Selection.ClearParagraphDirectFormatting
    
    Dim curPar As Paragraph
    Dim lastPar As Paragraph

    Dim firstParagraphInArticle As Paragraph
    Dim countParagraphsInArticle As Integer

      For Each curPar In ActiveDocument.Sections(2).Range.Paragraphs
        Dim curParText As String
        Dim i As Integer
        
        curPar.Range.Select
        Selection.ClearParagraphDirectFormatting
        
        If curPar.Style = "Überschrift 2" Then
            If Not lastPar Is Nothing Then
                lastPar.SpaceAfter = 6
            End If
        End If
        
        If curPar.Style = "Überschrift 1" Then
            curParText = curPar.Range.Text
            
            i = InStr(1, curParText, ". ")
            If i > 0 Then
                ActiveDocument.Range(curPar.Range.Start, curPar.Range.Start + i).Delete
            End If
        End If
        
        If curPar.Style = "Überschrift 2" Then
            ' MsgBox ("yes")
            ' MsgBox (curPar.Range.Text)
            
            curParText = curPar.Range.Text
            
            i = 0
            If InStr(1, curParText, "Art. ") = 1 Then
                i = i + Len("Art. ") + 1
                Dim c As Integer
                While IsNumeric(Mid(curParText, i, 1))
                    i = i + 1
                Wend
            End If
            If i > 0 Then
                ActiveDocument.Range(curPar.Range.Start, curPar.Range.Start + i).Delete
                ActiveDocument.Range(curPar.Range.Start).InsertBefore (" " & Chr(11))
            End If
        End If
        
        ' if there is only one legal paragraph in an article there should be no number
        If curPar.Style = "Überschrift 2" Then
            If FirstLegalParagraphWithoutNumber And Not (firstParagraphInArticle Is Nothing) And countParagraphsInArticle = 1 Then
                firstParagraphInArticle.Style = "Standard"
            End If
            Set firstParagraphInArticle = Nothing
            countParagraphsInArticle = 0
            
        End If
        If curPar.Style = "Scroll List Number" Or curPar.Style = "Standard" Then
            countParagraphsInArticle = countParagraphsInArticle + 1
            If firstParagraphInArticle Is Nothing Then
                Set firstParagraphInArticle = curPar
            End If
        End If
        
        Set lastPar = curPar
    Next
    
    If FirstLegalParagraphWithoutNumber And Not (firstParagraphInArticle Is Nothing) And countParagraphsInArticle = 1 Then
        firstParagraphInArticle.Style = "Standard"
    End If
End Sub
Sub FixIdentification()
    Dim shp As Shape
    Dim str As String

    For Each shp In ActiveDocument.Shapes
        ' Debug.Print (shp.Name)
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
            
            ' Reset formating
            shp.TextFrame.TextRange.Paragraphs(1).Style = "Inhaltssteuerelementtextbox"
        End If
    Next
End Sub

' Tables in Panels
Sub FixTables()
    Dim tbl As Table
    
    For Each tbl In ActiveDocument.Sections(2).Range.Tables
        If tbl.Style = "Scroll Table Normal Wide" Then
            ' Debug.Print (tbl.Style)
            tbl.Style = "Scroll Table Normal"
            tbl.PreferredWidthType = wdPreferredWidthPoints
            tbl.PreferredWidth = CentimetersToPoints(16)
            tbl.Rows.LeftIndent = tbl.Rows.LeftIndent - CentimetersToPoints(5.2)
        End If
    Next
End Sub

