VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const coniDateColumn        As Integer = 1 '�f�[�^���t��
Const coniStartRow          As Integer = 2 '�f�[�^�J�n�s
Const coniStartColumn       As Integer = 3 '�f�[�^�J�n��
Const coniSearchStartRow    As Integer = 1 '�����J�n�s
Const coniSearchColumn      As Integer = 1 '�����J�n��
Const consMsgTitle          As String = "����g�p�ʎ擾�}�N��" '���b�Z�[�W�_�C�A���O�̃^�C�g��
Const coniErrorMsgCol       As Integer = 17 '�G���[���b�Z�[�W��
Const coniCalcOffsetCol     As Integer = 13 '�v�Z�p�̕ύX�t���O��܂ł̃I�t�Z�b�g�񐔂��`
Const coniMaxRow            As Integer = 1000 '�ő又���s
Dim SubRowCnt               As Integer        '�ړ��p�f�[�^�s



'�w�b�_�ƃt�b�^�̓Ǎ� �I���t�H���_���̑S�u�b�N
Private Sub CommandButton1_Click()
    On Error GoTo Error_Handle
    Dim iRowCnt As Integer          '�s��
    Dim iAns    As Integer          '���b�Z�[�W�A���T�[
    Dim sMsg    As String           '���b�Z�[�W���e
    Dim sPath   As String           '�I���t�H���_
    Dim objFs   As Object           'FileSystemObject
    Dim objFld  As Object           '�t�H���_�z��
    Dim objFl   As Object           '�t�@�C��
    Dim objFo   As Object           '�t�H���_
    Dim SearchFileName As String    '�����Ώۂ̃t�H���_��
    
    SubRowCnt = 2
    
    sMsg = "�t�H���_��I�����Ă��������B" & vbCrLf & "�t�H���_���̑SExcel�t�@�C�����猟�����܂��B"
    iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)
    
    '�V�[�g�̏�����
    Call funSheetClear
    
    '�����Ώۂ̃t�H���_�����Z�b�g
    If OptionButton1.Value = True Then
        SearchFileName = "ke1nwnecz01_����g�p��.csv"
        ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = "ke1nwnecz01"
    Else
        SearchFileName = "ke2nwnecz01_����g�p��.csv"
        ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = "ke2nwnecz01"
    End If
    
    sPath = ""
    Call funSelectFolder(sPath)
    '�����I�΂Ȃ�����(�L�����Z��)�ꍇ
    If sPath = "" Then
        sMsg = "�������I�����܂��B"
        iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)
        Exit Sub
    End If
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    'sPath�z���̃t�H���_��z��Ƃ���objFld�ɑ��
    Set objFld = objFs.GetFolder(sPath)
    
    '�I���t�H���_���̃t�H���_�������J��
    Application.ScreenUpdating = False
    
    '�t�H���_�z��objFld����t�H���_objFo�����ԂɎ��o����funOpenSubFolder���Ăяo��
    'funOpenSubFolder�̈����̓t�H���_�̐�΃p�X�ƌ����Ώۃt�@�C����
    For Each objFo In objFld.SubFolders
        Call funOpenSubFolder(objFs.GetAbsolutePathName(objFo), SearchFileName)
    Next
    
    
    For Each objFl In objFld.Files
        If objFl.Name = SearchFileName Then
            Workbooks.Open Filename:=sPath + "\" + objFl.Name, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
            Call funReadValue
        End If
        
Return_E:
    Next
    Application.ScreenUpdating = True

    ' �I��
    sMsg = "�������I�����܂����B"
    iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)

Exit Sub

Error_Handle:
    '���Ƀt�@�C�����J���Ă��ē��e��j�������L�����Z�������ꍇ
    If Err.Number = 1004 Then
        sMsg = objFl.Name & "��ǂݍ��܂��A�����𑱂��܂��B"
        iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)
        Resume Return_E
    Else
        sMsg = "�\�����ʃG���[�ł��B" & vbCrLf _
             & "�G���[�ԍ��F" & Err.Number & vbCrLf _
             & "�\�[�X�F" & Err.Source & vbCrLf _
             & "�����F" & Err.Description
        iAns = MsgBox(sMsg, vbCritical + vbOKOnly, consMsgTitle)
    End If

End Sub

'�t�H���_�I��
Private Function funSelectFolder(ByRef rsPath As String)
    Dim objShell As Object  'Shell.Application�I�u�W�F�N�g
    Dim objPath As Object   '�I���t�H���_�̃p�X���i�[����I�u�W�F�N�g
    
    Set objShell = CreateObject("Shell.Application")
    Set objPath = objShell.BrowseForFolder(&O0, "�t�H���_��I��ł�������", &H1 + &H10, "")
    
    'objPath�I�u�W�F�N�g�̃p�X�v���p�e�B��rsPath�ɑ��
    If Not objPath Is Nothing Then
        rsPath = objPath.Items.Item.Path
    End If
    
    '�I�u�W�F�N�g�����
    Set objShell = Nothing
    Set objPath = Nothing
End Function

'�t�@�C�������o���v���V�[�W��
Private Function funOpenSubFolder(ByVal SubsPath As String, ByVal SearchFileName As String)
    
    'objFs��FileSystemObject���Z�b�g
    Set objFs = CreateObject("Scripting.FileSystemObject")
    
    '�����Ŏ󂯎�����p�X�ȉ��ɂ���y�t�H���_�z��z��Ƃ��Ď擾
    Set objFld = objFs.GetFolder(SubsPath)
    '�����Ŏ󂯎�����p�X�ȉ��ɂ���y�t�H���_�z�������J���Ă��̃v���V�[�W�����Ăяo���i�ċA�����j
    For Each objFo In objFld.SubFolders
        Call funOpenSubFolder(objFs.GetAbsolutePathName(objFo), SearchFileName)
    Next
    
    '�����Ŏ󂯎�����p�X�ȉ��ɂ���y�t�@�C���z�������J��
    For Each objFl In objFld.Files
        '�I���t�H���_���ɂ��錟�������Ɉ�v����Excel�t�@�C�����J���AfunReadValue���Ăяo��
        If objFl.Name = SearchFileName Then
            '[��΃p�X\��v�����t�@�C����]��Excel�t�@�C����ǂݎ���p�ŊJ��
            Workbooks.Open Filename:=SubsPath + "\" + objFl.Name, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
            Call funReadValue
        End If
    Next

End Function

'����g�p�ʂ�MAX,MIN,AVRAGE�𔲂��o����ThisWorkBook�ɓ]�L����v���V�[�W��
Private Function funReadValue()
    Dim MaxInRow  As Integer        'IN��Max�s
    Dim MinInRow  As Integer        'IN��Min�s
    Dim AvrInRow  As Integer        'IN��Avr�s
    Dim MaxOutRow As Integer        'OUT��Max�s
    Dim MinOutRow As Integer        'OUT��Min�s
    Dim AvrOutRow As Integer        'OUT��Avr�s
      
    Dim DateValue As String         '�e�l�̓��t
    
    Dim MaxInValue(29)  As String   'IN��Max�̒l���i�[����z��@29�|�[�g����̂�29�̌Œ蒷�z��
    Dim MinInValue(29)  As String   'IN��Min�̒l���i�[����z��@29�|�[�g����̂�29�̌Œ蒷�z��
    Dim AvrInValue(29)  As String   'IN��Avr�̒l���i�[����z��@29�|�[�g����̂�29�̌Œ蒷�z��
    Dim MaxOutValue(29) As String   'OUT��Max�̒l���i�[����z��@29�|�[�g����̂�29�̌Œ蒷�z��
    Dim MinOutValue(29) As String   'OUT��Min�̒l���i�[����z��@29�|�[�g����̂�29�̌Œ蒷�z��
    Dim AvrOutValue(29) As String   'OUT��Avr�̒l���i�[����z��@29�|�[�g����̂�29�̌Œ蒷�z��
    
    '1�s�ڂ���1000�s�ڂ܂�"�ő�"��������܂ő��j
    For startRow = 1 To coniMaxRow
        If "�ő�" = ActiveSheet.Cells(startRow, coniSearchColumn).Value Then
           MaxInRow = startRow
           MinInRow = startRow + 1
           AvrInRow = startRow + 2
           Exit For
        End If
    Next
   
   'IN�s�̏I��肩��1000�s�ڂ܂�"�ő�"��������܂ő��j
    For startRow = AvrInRow To coniMaxRow
        If "�ő�" = ActiveSheet.Cells(startRow, coniSearchColumn).Value Then
            MaxOutRow = startRow
            MinOutRow = startRow + 1
            AvrOutRow = startRow + 2
            Exit For
        End If
    Next
       
    'A3�s������t���擾
    DateValue = Cells(3, 1)
   
    Dim i As Integer    '�z�񑀍�ׂ̈̓Y��
    i = 0
    
    '�e�z��ɑS�|�[�g���̒l�����Ă���
    For startColumn = 3 To 31
        MaxInValue(i) = Cells(MaxInRow, startColumn)
        MinInValue(i) = Cells(MinInRow, startColumn)
        AvrInValue(i) = Cells(AvrInRow, startColumn)
        MaxOutValue(i) = Cells(MaxOutRow, startColumn)
        MinOutValue(i) = Cells(MinOutRow, startColumn)
        AvrOutValue(i) = Cells(AvrOutRow, startColumn)
        i = i + 1
    Next
    '�����Ώۃu�b�N�����
    ActiveWorkbook.Close
    
    
    Dim SH As Worksheet                                 '����ȉ���ThisWorkbook�ł̏������ݍ�Ƃ̂��߁uThisWorkbook�`�v��ϐ��ɓ����
    Set SH = ThisWorkbook.Worksheets("Sheet1")          '�iWith��ŏ����Ă������j
    
    Dim j As Integer '�z�񑀍�ׂ̈̓Y��
    j = 0
         
    'MaxIn�̏����o��
    For startColumn = 3 To 31
        SH.Cells(SubRowCnt, startColumn) = MaxInValue(j)
        j = j + 1
    Next
    '���t�̏�������
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    '����s�̈ړ�
    SubRowCnt = SubRowCnt + 1
    j = 0
     
    'MinIn
    For startColumn = 3 To 31
        SH.Cells(SubRowCnt, startColumn) = MinInValue(j)
        j = j + 1
    Next
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    SubRowCnt = SubRowCnt + 1
    j = 0
     
    'AvrIn
    For startColumn = 3 To 31
        SH.Cells(SubRowCnt, startColumn) = AvrInValue(j)
        j = j + 1
    Next
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    SubRowCnt = SubRowCnt + 1
    j = 0
     
    'MaxOut
    For startColumn = 3 To 31
       SH.Cells(SubRowCnt, startColumn) = MaxOutValue(j)
       j = j + 1
    Next
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    SubRowCnt = SubRowCnt + 1
    j = 0
     
    'MinOut
    For startColumn = 3 To 31
       SH.Cells(SubRowCnt, startColumn) = MinOutValue(j)
       j = j + 1
    Next
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    SubRowCnt = SubRowCnt + 1
    j = 0
     
    'AvrOut
    For startColumn = 3 To 31
       SH.Cells(SubRowCnt, startColumn) = AvrOutValue(j)
       j = j + 1
    Next
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    SubRowCnt = SubRowCnt + 1
    j = 0
     
End Function
'Sheet�̒l���N���A����v���V�[�W��
Private Function funSheetClear()
    Application.EnableEvents = False
    '�Ώۋ@���\������Z�����N���A
    ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = " "
    '���ʒl��\������Z�����N���A
    ThisWorkbook.Worksheets("Sheet1").Range(ThisWorkbook.Worksheets("Sheet1").Cells(coniStartRow, coniStartColumn), ThisWorkbook.Worksheets("Sheet1").Cells(coniStartRow, coniStartColumn).SpecialCells(xlLastCell)).ClearContents
    '���t��\������Z�����N���A
    ThisWorkbook.Worksheets("Sheet1").Range(ThisWorkbook.Worksheets("Sheet1").Cells(coniStartRow, coniDateColumn), ThisWorkbook.Worksheets("Sheet1").Cells(coniMaxRow, coniDateColumn)).ClearContents
    Application.EnableEvents = True
End Function
