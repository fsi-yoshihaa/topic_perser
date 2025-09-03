Attribute VB_Name = "DumpLog"

Sub DumpLog()
    ' �ǂݍ��݃t�@�C���̃p�X���擾
    Dim inputFile As String
    inputFile = Range("B2")
    ' �o�͐�t�H���_�̃p�X���擾
    Dim outputFolderPath As String
    outputFolderPath = Range("B5")
    ' �󔒂��폜����
    inputFile = Replace(inputFile, " ", "")
    inputFile = Replace(inputFile, "�@", "")
    
     ' ���͂�����Ă��Ȃ��A�w�肳�ꂽ�t�@�C�������݂��Ȃ��ꍇ�̏���
    If inputFile = "" Or Dir(inputFile) = "" Then
        MsgBox "�ǂݍ��݃t�@�C�������݂��܂���B", vbExclamation
        '�Z���Ƀt�H�[�J�X���ړ�����
        Range("B2").Select
        Exit Sub
    End If
    
    If outputFolderPath = "" Then ' �o�͐�t�H���_���󔒂̏ꍇ
        ' �ǂݍ��݃t�@�C���Ɠ����t�H���_���o�͐�t�H���_�Ƃ���
        outputFolderPath = Left(inputFile, (InStrRev(inputFile, "\") - 1))
        Range("B5") = outputFolderPath
    End If
    
    Dim inputFileLen As Integer ' �ǂݍ��݃t�@�C���̃p�X�̒���
    inputFileLen = Len(inputFile)
    Dim outputFilePath As String ' �o�̓t�@�C���̃p�X
    Dim outputFileName As String ' �o�̓t�@�C����
    outputFileName = makeFileName()
    outputFilePath = outputFolderPath & "\" & outputFileName
    Dim outputSheetName As String ' �o�͂���V�[�g��
    outputSheetName = "�_���v"
    Dim fOutputObj As Object
    Dim outputWb As Workbook  ' �o�͂��郏�[�N�u�b�N
    Dim outputWs As Worksheet ' �o�͂���V�[�g
    Dim tempFilePath As String '�@�V�����o�̓t�H���_�[
    
    If Dir(inputFile) <> "" Then ' B2�Ŏw�肳�ꂽ�ǂݍ��݃t�@�C�������݂���ꍇ
        If Dir(outputFolderPath, vbDirectory) <> "" Then ' B5�Ŏw�肳�ꂽ�o�͐�t�H���_�����݂���ꍇ
            ' �u�b�N��V�K�쐬
            Set outputWb = Workbooks.Add
            ' �V�[�g����ύX
            Set outputWs = outputWb.Sheets(1)
            outputWs.Name = outputSheetName
            
             '�@�V�����o�̓t�H���_�[
            tempFilePath = outputFolderPath & "\Temp_CRLF.txt"
            ' ���t�@�C�����R�s�[���ĉ��s�R�[�h�ϊ�
            Call LfToCrlfCopy(inputFile, tempFilePath)
    
            ' �_���v�����i���t�@�C���ł͂Ȃ��ꎞ�t�@�C�����g�p�j
            Call OutputDumpData(tempFilePath, outputWs)

            ' �u�b�N��ۑ�
            outputWb.SaveAs outputFilePath

            ' �u�b�N�����
            outputWb.Close
            
            ' �ꎞ�t�@�C�����폜
            If Dir(tempFilePath) <> "" Then
                 Kill tempFilePath
            End If

            
             ' ����ɏo�͂��ꂽ���Ƃ��������b�Z�[�W��\��!!!!!!!'
            Call ShowCompletionMessage
            
        Else
            ' �G���[���b�Z�[�W���o��
            MsgBox "�o�͐�t�H���_��������܂���", vbExclamation
        End If
        
    Else  ' B2�Ŏw�肳�ꂽ�ǂݍ��݃t�@�C�������݂��Ȃ��ꍇ
        ' �G���[���b�Z�[�W���o��
        MsgBox "�ǂݍ��݃t�@�C����������܂���", vbExclamation
        
    End If
    
End Sub


Function makeFileName() As String
    ' ���t_�������擾
    Dim dateTime
    dateTime = Now()
    
    ' ������ɕϊ�
    Dim retStr As String
    retStr = Format(dateTime, "yyyymmdd_hmmss")
    
    makeFileName = retStr & "_log_dump.xlsx"
End Function

'#### �o�� ####'
Function OutputDumpData(ByVal tempFilePath As String, ByVal outputWs As Worksheet)
    ' �o�͐�̃V�[�g���N���A����
    outputWs.Cells.Clear
    
    ' �ϐ��錾
    Dim serchStr As String              ' �o�͂���f�[�^���𔻒f����ڈ�
    Dim flgRead As Boolean              ' �ǂݍ��݃f�[�^�t���O
    Dim readLine As String              ' �ǂݍ��񂾗�̕�����
    Dim headerList As New Collection    ' �w�b�_�̕�����̃��X�g
    Dim lastKey As String               ' �Ō��key�i�w�b�_�̍��ځj
    Dim clmMax As Integer               ' �ő��
    Dim row As Integer                  ' �s
    Dim rowHeader As Integer            ' �w�b�_���o�͂���s
    Dim clm As Integer                  ' ��
    Dim clmNo As Integer                ' No��\�������
    Dim cntBlock As Integer             ' �o�̓u���b�N�̃J�E���g
    Dim cntData As Integer              ' �f�[�^�̃J�E���g
    
    serchStr = "---"
    flgRead = False
    cntBlock = 0
    row = 3         ' ���s�ڂ���o�͂��邩��ݒ�
    rowHeader = row
    clmNo = 2       ' ����ڂ���o�͂��邩��ݒ�
    cntData = 1
    clmMax = clmNo
    
    ' ���̓t�@�C���I�[�v��
    Open tempFilePath For Input As #1
    
    Do Until EOF(1)
        ' 1�񂸂ǂݍ���
        Line Input #1, readLine
        
        If readLine = serchStr Then ' --- �̏ꍇ
            ' �ǂݍ��݃t���O��ON�ɂ���
            flgRead = True
            
            cntBlock = cntBlock + 1
            cntData = 1
            row = row + 1
            clm = clmNo + 1
            
            If cntBlock = 1 Then ' 1�ڂ̃u���b�N�̏ꍇ
                ' [No](�w�b�_)���o��
                outputWs.Cells(rowHeader, clmNo).NumberFormatLocal = "@"
                outputWs.Cells(rowHeader, clmNo) = "No"
            End If
            
            ' No����o��
            outputWs.Cells(row, clmNo) = cntBlock
            
        ElseIf flgRead Then ' [---] �ȍ~�̏ꍇ
            Dim key As String
            Dim item As String
            
            If Left(readLine, 1) = "-" Then ' -����n�܂镶����̏ꍇ
            
                If cntBlock = 1 Then ' 1�ڂ̃u���b�N�̏ꍇ
                    ' �w�b�_���o��
                    outputWs.Cells(rowHeader, clm).NumberFormatLocal = "@"
                    outputWs.Cells(rowHeader, clm) = lastKey & "_" & cntData
                End If
                
                ' �A�C�e�����o��
                item = Replace(readLine, " ", "")
                item = Replace(item, "�@", "")
                outputWs.Cells(row, clm).NumberFormatLocal = "@"
                outputWs.Cells(row, clm) = item
                
                If clm > clmMax Then ' �o�͂����񂪍ő�񐔂�葽���ꍇ
                    clmMax = clm
                End If
                clm = clm + 1
                cntData = cntData + 1
                
            ElseIf InStr(readLine, ":") > 0 Then '[:]���܂ޕ�����̏ꍇ
                item = GetItemStr(readLine)
                key = GetKeyStr(readLine)
                lastKey = key
                cntData = 1
                
                If cntBlock = 1 Then ' 1�ڂ̃u���b�N�̏ꍇ
                    ' �w�b�_���o��
                    outputWs.Cells(rowHeader, clm).NumberFormatLocal = "@"
                    outputWs.Cells(rowHeader, clm) = key
                End If
                ' �A�C�e�����o��
                outputWs.Cells(row, clm).NumberFormatLocal = "@"
                outputWs.Cells(row, clm) = item
                
                If clm > clmMax Then  ' �o�͂����񂪍ő�񐔂�葽���ꍇ
                    clmMax = clm
                End If
                clm = clm + 1
            End If
        End If
    Loop
    
    ' �g��ǉ�
    outputWs.Range(outputWs.Cells(rowHeader, clmNo), outputWs.Cells(row, clmMax)).Borders.LineStyle = xlContinuous
    ' �񕝒���
    outputWs.Columns.AutoFit
    
    ' �t�@�C���N���[�Y
    Close #1
    
End Function


'#### Key���擾 ####'
Function GetKeyStr(ByVal inputStr As String) As String
    Dim retStr As String
    retStr = Left(inputStr, InStr(inputStr, ":") - 1)
    '�󔒕�����͍폜����
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "�@", "")
    GetKeyStr = retStr
End Function


'#### item���擾 ####'
Function GetItemStr(ByVal inputStr As String) As String
    Dim retStr As String
    retStr = Mid(inputStr, InStr(inputStr, ":") + 1)
    '�󔒕�����͍폜����
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "�@", "")
    GetItemStr = retStr
End Function


'### ����ɏo�͂ł������\��###'
Sub ShowCompletionMessage()
    Dim res As Integer
    res = MsgBox("�������܂���", vbOKOnly)
End Sub


'#### ���s�R�[�h�ϊ� ####'

Function LfToCrlfCopy(ByVal inputFile As String, ByVal tempFilePath As String)
    Dim FileNum As Integer
    Dim FileContent As String
    Dim NewContent As String

    ' ���t�@�C����ǂݍ���
    FileNum = FreeFile
    Open inputFile For Input As #FileNum
    FileContent = Input(LOF(FileNum), #FileNum)
    Close #FileNum

    ' ���s�R�[�h��ϊ��iLF �� CRLF�j
    NewContent = Replace(FileContent, vbLf, vbCrLf)

    ' �ꎞ�t�@�C���ɕۑ�
    FileNum = FreeFile
    Open tempFilePath For Output As #FileNum
    Print #FileNum, NewContent
    Close #FileNum
End Function

