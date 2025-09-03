Attribute VB_Name = "GetFileName"
Sub GetinputFile()

    ' �t�@�C�������擾
    Dim inputFile
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = False Then
            Exit Sub
        End If
        inputFile = .SelectedItems(1)
    End With
    
    ' �t�@�C�������o��
    Range("B2") = inputFile
End Sub
