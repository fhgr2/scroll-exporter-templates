Attribute VB_Name = "Regulation_Macros"
Sub ExitReadingLayout()

    If IsExported() Then
        If ShouldRunOnceAfterExport() Then
        
            progressBar.Show (vbModeless)
            
            ' Common
            progressBar.tasksTextBox.text = "Allgemeine Korrekturen"
            DoEvents
            
            Call FixAllPlaceholdersInHeadersFooters("Inhaltssteuerelementtextbox")
            Call FixPlaceholders(2, "Inhaltssteuerelementtextbox")
            Call FixBold(4)
            Call SetDocumentPropertiesFromShapeContents
            ' Call FixTableOfContents
            
            ' specific
            progressBar.tasksTextBox.text = "Reglements-spezifische Korrekturen"
            DoEvents
            
            progressBar.tasksTextBox.text = "Formatierung von Artikeln korrigieren... "
            DoEvents
            Call FixArticles
            
            Call FixTables
            Call FixGestuetztAuf
            Call RemoveEmptyBasis
            
            progressBar.tasksTextBox.text = "Layout korrigieren... "
            DoEvents
            Call FixSchusterjungen
            
            SetRun (True)
            
            progressBar.Hide
        Else
            SetRun (False)
        End If
    End If

End Sub

Sub FixSchusterjungen()
    Dim count As Integer
    count = 0
    Do While (FixSchusterjunge() And count < 10)
        count = count + 1
    Loop
    Debug.Print (count)
    
End Sub

Function FixSchusterjunge() As Boolean
    Dim curPar As Paragraph
    Dim lastPar As Paragraph
    Dim firstParagraphInArticle As Paragraph
    Dim firstArticleInChapter As Paragraph
    Dim curArticle As Paragraph
    Dim curChapter As Paragraph
    Dim countParagraphsInArticle As Integer
    Dim countArticlesInChapter As Integer
    Dim pageChapter As Integer
    Dim pageArticle As Integer
    Dim pageFirstParagraph As Integer
    Dim pageSecondParagraph As Integer
    
    pageFirstParagraph = 0
    pageSecondParagraph = 0
    
    For Each curPar In ActiveDocument.Sections(4).Range.Paragraphs
        Dim curParText As String
        
        ' Debug.Print (curPar.Range.Information(wdActiveEndAdjustedPageNumber))
        ' Debug.Print (curPar.Range.text)
        
        If curPar.style = "Überschrift 1" Then
            ' before
        
            ' update chapter
            Set curChapter = curPar
            pageChapter = curPar.Range.Information(wdActiveEndAdjustedPageNumber)
            ' update article
            Set firstArticleInChapter = Nothing
            countArticlesInChapter = 0
            
            ' after
        End If
        
        If curPar.style = "Überschrift 2" Then
            ' before
            
            ' fix article
            Set curArticle = curPar
            If (countArticlesInChapter = 0) Then
                Set firstArticleInChapter = curArticle
            End If
            countArticlesInChapter = countArticlesInChapter + 1
            pageArticle = curPar.Range.Information(wdActiveEndAdjustedPageNumber)
            ' fix paragraph
            pageFirstParagraph = 0
            pageSecondParagraph = 0
            countParagraphsInArticle = 0
            Set firstParagraphInArticle = Nothing
            
            ' after
            If (countArticlesInChapter = 1) Then
                pageFirstArticle = curPar.Range.Information(wdActiveEndAdjustedPageNumber)
                If (pageFirstArticle > pageChapter) Then
                    curChapter.Range.Select
                    Selection.ParagraphFormat.PageBreakBefore = True
                    FixSchusterjunge = True
                    Exit Function
                End If
            End If
        
        
        End If
        
                
        If curPar.style = "Scroll List Number" Or curPar.style = "Standard" Then
            ' before
            
            ' fix paragraph
            countParagraphsInArticle = countParagraphsInArticle + 1
            If (countParagraphsInArticle = 1) Then
                Set firstParagraphInArticle = curPar
                pageFirstParagraph = curPar.Range.Information(wdActiveEndAdjustedPageNumber)
            End If
            If (countParagraphsInArticle = 2) Then
                pageSecondParagraph = curPar.Range.Information(wdActiveEndAdjustedPageNumber)
            End If
                
            ' after
            If (countParagraphsInArticle = 1) Then
                If (pageArticle < pageFirstParagraph) Then
                    If Not (curArticle Is Nothing) Then
                        curArticle.Range.Select
                        Selection.ParagraphFormat.PageBreakBefore = True
                        FixSchusterjunge = True
                        Exit Function
                    End If
                End If
            End If
            
            If (countParagraphsInArticle = 2) Then
                If (pageSecondParagraph > pageFirstParagraph) Then
                    Debug.Print ("Schusterjunge")
                    ' Debug.Print (curArticle.Range.text)
                    ' curArticle.Range.InsertBefore ("X")
                    If (countArticlesInChapter = 1) Then
                        curChapter.Range.Select
                    Else
                        curArticle.Range.Select
                    End If
                    Selection.ParagraphFormat.PageBreakBefore = True
                    FixSchusterjunge = True
                    Exit Function
                End If
            End If
        End If
        
        Set lastPar = curPar
    Next
    FixSchusterjunge = False

End Function

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
        tbl.Select
        Selection.ParagraphFormat.KeepWithNext = True

    Next
End Sub

Sub RemoveEmptyBasis()
    Dim oShape As Shape
    Dim text As String
    
    Set oShape = GetShape("basis")
    If oShape Is Nothing Then
        Exit Sub
    End If
    text = oShape.TextFrame.TextRange.text
    If Len(text) > 5 Then
        Exit Sub
    End If
    
    ' Remove
    Dim oRange As Range
    Set oRange = ActiveDocument.Bookmarks("basis").Range
    If oRange Is Nothing Then
        Exit Sub
    End If
    oRange.Delete
End Sub


