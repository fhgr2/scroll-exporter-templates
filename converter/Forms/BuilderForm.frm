VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuilderForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2136
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4332
   OleObjectBlob   =   "BuilderForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "BuilderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chooseButton_Click()
    Dim dlgSaveAs As FileDialog
    Set dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
    dlgSaveAs.Show
    Debug.Print (dlgSaveAs.SelectedItems(1))
    templateTextBox.Text = dlgSaveAs.SelectedItems(1)
End Sub

Private Sub closeButton_Click()
    ' cancelled = True
    BuilderForm.Hide
    
End Sub

Private Sub convertButton_Click()
    Dim doc As Word.Document
    Dim name As String
    
    If (convertPagePropertiesCheckBox.Value) Then
        Set doc = Documents.Open(templateTextBox.Text)
        name = doc.Path + Application.PathSeparator + "PageProperties.docx"
        Call doc.SaveAs2(name)
        Call ConvertToPageProperties(doc)
        Call doc.Close(True)
    End If
    
    If (convertConfluenceCheckBox.Value) Then
        Set doc = Documents.Open(templateTextBox.Text)
        name = doc.Path + Application.PathSeparator + "Confluence.docx"
        Call doc.SaveAs2(name)
        Call ConvertToConfluence(doc)
        Call doc.Close(True)
    End If
    
    If (convertStandardCheckBox.Value) Then
        Set doc = Documents.Open(templateTextBox.Text)
        name = doc.Path + Application.PathSeparator + "Standard.docx"
        Call doc.SaveAs2(name)
        Call ConvertToStandard(doc)
        Call doc.Close(True)
    End If
    
End Sub

Private Sub convertStandardCheckBox_Click()

End Sub
