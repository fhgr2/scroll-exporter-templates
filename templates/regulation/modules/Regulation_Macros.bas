Attribute VB_Name = "Regulation_Macros"
Sub ExitReadingLayout()
    If IsExported() Then
        If ShouldRunOnceAfterExport() Then
            
            ' Common
            Call FixAllPlaceholdersInHeadersFooters("Inhaltssteuerelementtextbox")
            Call FixPlaceholders(2, "Inhaltssteuerelementtextbox")
            Call FixBold(4)
            Call SetDocumentPropertiesFromShapeContents
            ' Call FixTableOfContents
            
            ' specific
            Call FixArticles
            Call FixTables
            Call FixGestuetztAuf
            
            SetRun (True)
        Else
            SetRun (False)
        End If
    End If

End Sub

Sub FixArticles()
    
    Const FirstLegalParagraphWithoutNumber As Boolean = False
    
    ' Selection.WholeStory
    ' Selection.ClearParagraphDirectFormatting
    
    Dim curPar As Paragraph
    Dim lastPar As Paragraph

    Dim firstParagraphInArticle As Paragraph
    Dim countParagraphsInArticle As Integer

      For Each curPar In ActiveDocument.Sections(4).Range.Paragraphs
        Dim curParText As String
        Dim i As Integer
        
        curPar.Range.Select
        Selection.ClearParagraphDirectFormatting
        
        If curPar.style = "Überschrift 2" Then
            If Not lastPar Is Nothing Then
                lastPar.SpaceAfter = 6
            End If
        End If
        
        If curPar.style = "Überschrift 1" Then
            curParText = curPar.Range.text
            
            i = InStr(1, curParText, ". ")
            If i > 0 Then
                ActiveDocument.Range(curPar.Range.Start, curPar.Range.Start + i).Delete
            End If
        End If
        
        If curPar.style = "Überschrift 2" Then
            ' MsgBox ("yes")
            ' MsgBox (curPar.Range.Text)
            
            curParText = curPar.Range.text
            
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
        If curPar.style = "Überschrift 2" Then
            If FirstLegalParagraphWithoutNumber And Not (firstParagraphInArticle Is Nothing) And countParagraphsInArticle = 1 Then
                firstParagraphInArticle.style = "Standard"
            End If
            Set firstParagraphInArticle = Nothing
            countParagraphsInArticle = 0
            
        End If
        If curPar.style = "Scroll List Number" Or curPar.style = "Standard" Then
            countParagraphsInArticle = countParagraphsInArticle + 1
            If firstParagraphInArticle Is Nothing Then
                Set firstParagraphInArticle = curPar
            End If
        End If
        
        Set lastPar = curPar
    Next
    
    If FirstLegalParagraphWithoutNumber And Not (firstParagraphInArticle Is Nothing) And countParagraphsInArticle = 1 Then
        firstParagraphInArticle.style = "Standard"
    End If
End Sub


Sub FixGestuetztAuf()
    Call FixPlaceholders(3, "Standard")
End Sub

' Tables in Panels
Sub FixTables()
    Dim tbl As Table
    
    For Each tbl In ActiveDocument.Sections(4).Range.Tables
        If tbl.style = "Scroll Table Normal Wide" Then
            ' Debug.Print (tbl.Style)
            tbl.style = "Scroll Table Normal"
            tbl.PreferredWidthType = wdPreferredWidthPoints
            tbl.PreferredWidth = CentimetersToPoints(16)
            tbl.Rows.LeftIndent = tbl.Rows.LeftIndent - CentimetersToPoints(5.2)
        Else
            tbl.AutoFitBehavior (wdAutoFitWindow)
        End If
    Next
End Sub

