Attribute VB_Name = "Macros"
Sub ExitReadingLayout()
    If IsExported() Then
        If ShouldRunOnceAfterExport() Then
        
            ' Common
            Call progressBar.Show(vbModeless)

            progressBar.tasksTextBox.text = "Allgemeine Korrekturen..."
            DoEvents
            
            Call FixAllPlaceholdersInHeadersFooters("Inhaltssteuerelementtextbox")
            Call FixPlaceholders(2, "Inhaltssteuerelementtextbox")
            Call FixBold(4)
            Call SetDocumentPropertiesFromShapeContents
            Call FixTableOfContents
            
            progressBar.tasksTextBox.text = "Spezifische Korrekturen..."
            DoEvents
        
            Call FixJIRAMacroExport
            
            SetRun (True)
        
            progressBar.Hide
        
        Else
            SetRun (False)
        End If
    End If
End Sub


Sub FixJIRAMacroExport()
    Call FixJIRAMacroExportInSection(5)
End Sub


Sub FixJIRAMacroExportInSection(section As Integer)
    Call FixJIRAMacroExportInRange(ActiveDocument.Sections(section).Range)
End Sub

Sub FixJIRAMacroExportInRange(rangeObj As Range)
    Dim shp As Shape
    Dim str As String
    Dim oTables As Tables
    Dim oTable As Table
    Dim oCello As Cell
    
    
    str = rangeObj.text
    

    For Each shp In rangeObj.ShapeRange
        ' only fix text boxes
        If (shp.Type = MsoShapeType.msoTextBox) Then
            
            Set oTables = shp.TextFrame.TextRange.Tables
            If oTables.Count > 0 Then
                
                Set oTable = shp.TextFrame.TextRange.Tables.Item(1)
                
                
                str = oTable.Range.text
                
                If oTable.Columns.Count() = 1 Then
                
                
                   Set oCello = oTable.Cell(0, 0)
                   
                   str = oCello.Range.text
                   
                      ' Trim white space from begin
                   i = 1
                   If i <= Len(str) Then
                       c = Mid(str, i, 1)
                   End If
                   
                   While i <= Len(str) And (c = Chr(13) Or c = Chr(32))
                       i = i + 1
                       c = Mid(str, i, 1)
                   Wend
                   str = Mid(str, i, Len(str) - i)
                   
                   ' trim white space from end
                   i = Len(str)
                   If i > 0 Then
                       c = Mid(str, i, 1)
                   End If
                   While i > 0 And (c = Chr(13) Or c = Chr(32))
                       i = i - 1
                       c = Mid(str, i, 1)
                   Wend
                
                   
                   str = Mid(str, 1, i)
                       
                   
                   
                   shp.TextFrame.DeleteText
                   shp.TextFrame.TextRange.text = str
                
                Else
                    ' Only Table
                    shp.TextFrame.TextRange.Tables.Item(1).Select
                    Selection.Copy
                                        
                    shp.TextFrame.DeleteText
                    Selection.Paste
                    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                    Selection.ParagraphFormat.LineSpacing = 0.7
                    Selection.ParagraphFormat.SpaceBefore = 0
                    Selection.ParagraphFormat.SpaceAfter = 0
                    
                    
                    
                    ' shp.TextFrame.TextRange.Copy
                    
                    
                End If
                
                'shp.TextFrame.
                
                ' shp.Select
                ' str = Selection.ShapeRange.TextFrame.TextRange.text
                
                ' Only first line
                ' str = Split(str, Chr(13))(0)
                ' Debug.Print (str)
        
                ' shp.TextFrame.TextRange.Delete
                ' shp.TextFrame.DeleteText
                ' shp.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                ' shp.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
                ' shp.TextFrame.TextRange.text = str
                
                ' set formating
                ' shp.TextFrame.TextRange.Paragraphs(1).style = style
            End If
        End If
    Next
End Sub

