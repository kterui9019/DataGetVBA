VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const coniDateColumn        As Integer = 1 'データ日付列
Const coniStartRow          As Integer = 2 'データ開始行
Const coniStartColumn       As Integer = 3 'データ開始列
Const coniSearchStartRow    As Integer = 1 '検索開始行
Const coniSearchColumn      As Integer = 1 '検索開始列
Const consMsgTitle          As String = "回線使用量取得マクロ" 'メッセージダイアログのタイトル
Const coniErrorMsgCol       As Integer = 17 'エラーメッセージ列
Const coniCalcOffsetCol     As Integer = 13 '計算用の変更フラグ列までのオフセット列数を定義
Const coniMaxRow            As Integer = 1000 '最大処理行
Dim SubRowCnt               As Integer        '移動用データ行



'ヘッダとフッタの読込 選択フォルダ内の全ブック
Private Sub CommandButton1_Click()
    On Error GoTo Error_Handle
    Dim iRowCnt As Integer          '行数
    Dim iAns    As Integer          'メッセージアンサー
    Dim sMsg    As String           'メッセージ内容
    Dim sPath   As String           '選択フォルダ
    Dim objFs   As Object           'FileSystemObject
    Dim objFld  As Object           'フォルダ配列
    Dim objFl   As Object           'ファイル
    Dim objFo   As Object           'フォルダ
    Dim SearchFileName As String    '検索対象のフォルダ名
    
    SubRowCnt = 2
    
    sMsg = "フォルダを選択してください。" & vbCrLf & "フォルダ内の全Excelファイルから検索します。"
    iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)
    
    'シートの初期化
    Call funSheetClear
    
    '検索対象のフォルダ名をセット
    If OptionButton1.Value = True Then
        SearchFileName = "ke1nwnecz01_回線使用量.csv"
        ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = "ke1nwnecz01"
    Else
        SearchFileName = "ke2nwnecz01_回線使用量.csv"
        ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = "ke2nwnecz01"
    End If
    
    sPath = ""
    Call funSelectFolder(sPath)
    '何も選ばなかった(キャンセル)場合
    If sPath = "" Then
        sMsg = "処理を終了します。"
        iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)
        Exit Sub
    End If
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    'sPath配下のフォルダを配列としてobjFldに代入
    Set objFld = objFs.GetFolder(sPath)
    
    '選択フォルダ内のフォルダを順次開く
    Application.ScreenUpdating = False
    
    'フォルダ配列objFldからフォルダobjFoを順番に取り出してfunOpenSubFolderを呼び出す
    'funOpenSubFolderの引数はフォルダの絶対パスと検索対象ファイル名
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

    ' 終了
    sMsg = "処理が終了しました。"
    iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)

Exit Sub

Error_Handle:
    '既にファイルを開いていて内容を破棄せずキャンセルした場合
    If Err.Number = 1004 Then
        sMsg = objFl.Name & "を読み込まず、処理を続けます。"
        iAns = MsgBox(sMsg, vbInformation + vbOKOnly, consMsgTitle)
        Resume Return_E
    Else
        sMsg = "予期せぬエラーです。" & vbCrLf _
             & "エラー番号：" & Err.Number & vbCrLf _
             & "ソース：" & Err.Source & vbCrLf _
             & "説明：" & Err.Description
        iAns = MsgBox(sMsg, vbCritical + vbOKOnly, consMsgTitle)
    End If

End Sub

'フォルダ選択
Private Function funSelectFolder(ByRef rsPath As String)
    Dim objShell As Object  'Shell.Applicationオブジェクト
    Dim objPath As Object   '選択フォルダのパスを格納するオブジェクト
    
    Set objShell = CreateObject("Shell.Application")
    Set objPath = objShell.BrowseForFolder(&O0, "フォルダを選んでください", &H1 + &H10, "")
    
    'objPathオブジェクトのパスプロパティをrsPathに代入
    If Not objPath Is Nothing Then
        rsPath = objPath.Items.Item.Path
    End If
    
    'オブジェクトを閉じる
    Set objShell = Nothing
    Set objPath = Nothing
End Function

'ファイルを取り出すプロシージャ
Private Function funOpenSubFolder(ByVal SubsPath As String, ByVal SearchFileName As String)
    
    'objFsにFileSystemObjectをセット
    Set objFs = CreateObject("Scripting.FileSystemObject")
    
    '引数で受け取ったパス以下にある【フォルダ】を配列として取得
    Set objFld = objFs.GetFolder(SubsPath)
    '引数で受け取ったパス以下にある【フォルダ】を順次開いてこのプロシージャを呼び出す（再帰処理）
    For Each objFo In objFld.SubFolders
        Call funOpenSubFolder(objFs.GetAbsolutePathName(objFo), SearchFileName)
    Next
    
    '引数で受け取ったパス以下にある【ファイル】を順次開く
    For Each objFl In objFld.Files
        '選択フォルダ内にある検索条件に一致したExcelファイルを開き、funReadValueを呼び出す
        If objFl.Name = SearchFileName Then
            '[絶対パス\一致したファイル名]のExcelファイルを読み取り専用で開く
            Workbooks.Open Filename:=SubsPath + "\" + objFl.Name, ReadOnly:=True, IgnoreReadOnlyRecommended:=True
            Call funReadValue
        End If
    Next

End Function

'回線使用量のMAX,MIN,AVRAGEを抜き出してThisWorkBookに転記するプロシージャ
Private Function funReadValue()
    Dim MaxInRow  As Integer        'INのMax行
    Dim MinInRow  As Integer        'INのMin行
    Dim AvrInRow  As Integer        'INのAvr行
    Dim MaxOutRow As Integer        'OUTのMax行
    Dim MinOutRow As Integer        'OUTのMin行
    Dim AvrOutRow As Integer        'OUTのAvr行
      
    Dim DateValue As String         '各値の日付
    
    Dim MaxInValue(29)  As String   'INのMaxの値を格納する配列　29ポートあるので29個の固定長配列
    Dim MinInValue(29)  As String   'INのMinの値を格納する配列　29ポートあるので29個の固定長配列
    Dim AvrInValue(29)  As String   'INのAvrの値を格納する配列　29ポートあるので29個の固定長配列
    Dim MaxOutValue(29) As String   'OUTのMaxの値を格納する配列　29ポートあるので29個の固定長配列
    Dim MinOutValue(29) As String   'OUTのMinの値を格納する配列　29ポートあるので29個の固定長配列
    Dim AvrOutValue(29) As String   'OUTのAvrの値を格納する配列　29ポートあるので29個の固定長配列
    
    '1行目から1000行目まで"最大"が見つかるまで走破
    For startRow = 1 To coniMaxRow
        If "最大" = ActiveSheet.Cells(startRow, coniSearchColumn).Value Then
           MaxInRow = startRow
           MinInRow = startRow + 1
           AvrInRow = startRow + 2
           Exit For
        End If
    Next
   
   'IN行の終わりから1000行目まで"最大"が見つかるまで走破
    For startRow = AvrInRow To coniMaxRow
        If "最大" = ActiveSheet.Cells(startRow, coniSearchColumn).Value Then
            MaxOutRow = startRow
            MinOutRow = startRow + 1
            AvrOutRow = startRow + 2
            Exit For
        End If
    Next
       
    'A3行から日付を取得
    DateValue = Cells(3, 1)
   
    Dim i As Integer    '配列操作の為の添字
    i = 0
    
    '各配列に全ポート分の値を入れていく
    For startColumn = 3 To 31
        MaxInValue(i) = Cells(MaxInRow, startColumn)
        MinInValue(i) = Cells(MinInRow, startColumn)
        AvrInValue(i) = Cells(AvrInRow, startColumn)
        MaxOutValue(i) = Cells(MaxOutRow, startColumn)
        MinOutValue(i) = Cells(MinOutRow, startColumn)
        AvrOutValue(i) = Cells(AvrOutRow, startColumn)
        i = i + 1
    Next
    '検索対象ブックを閉じる
    ActiveWorkbook.Close
    
    
    Dim SH As Worksheet                                 'これ以下はThisWorkbookでの書き込み作業のため「ThisWorkbook～」を変数に入れる
    Set SH = ThisWorkbook.Worksheets("Sheet1")          '（With句で書いてもいい）
    
    Dim j As Integer '配列操作の為の添字
    j = 0
         
    'MaxInの書き出し
    For startColumn = 3 To 31
        SH.Cells(SubRowCnt, startColumn) = MaxInValue(j)
        j = j + 1
    Next
    '日付の書きだし
    SH.Cells(SubRowCnt, coniDateColumn) = DateValue
    '操作行の移動
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
'Sheetの値をクリアするプロシージャ
Private Function funSheetClear()
    Application.EnableEvents = False
    '対象機器を表示するセルをクリア
    ThisWorkbook.Worksheets("Sheet1").Cells(1, 1) = " "
    '結果値を表示するセルをクリア
    ThisWorkbook.Worksheets("Sheet1").Range(ThisWorkbook.Worksheets("Sheet1").Cells(coniStartRow, coniStartColumn), ThisWorkbook.Worksheets("Sheet1").Cells(coniStartRow, coniStartColumn).SpecialCells(xlLastCell)).ClearContents
    '日付を表示するセルをクリア
    ThisWorkbook.Worksheets("Sheet1").Range(ThisWorkbook.Worksheets("Sheet1").Cells(coniStartRow, coniDateColumn), ThisWorkbook.Worksheets("Sheet1").Cells(coniMaxRow, coniDateColumn)).ClearContents
    Application.EnableEvents = True
End Function
