Attribute VB_Name = "Module1"
Sub GetOutputFolder()
    ' �t�H���_�����擾
    Dim outputFolder As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then ' �L�����Z���{�^��������
            Exit Sub
        End If
        outputFolder = .SelectedItems(1)
    End With
    
    ' �t�@�C�������o��
    Range("B5") = outputFolder
End Sub
