'----------------------------------------------------------------------------------------------------+
'-------------------------------------------*使用方法*-----------------------------------------------+
'----------------------------------------------------------------------------------------------------+
'各モジュールのトップステートメントに以下の定数を記述
'   Private Const MODULE_NAME As String = "＜モジュール名＞"
'
'各プロシージャの構成を以下のものにする
'・親プロシージャの場合
'Const PROCEDURE_NAME As String = "＜プロシージャ名＞"
'Call ErrorLog.Initialize
'Call ErrorLog.TraceListPush(MODULE_NAME, PROCEDURE_NAME)
'On Error GoTo ErrorHandler
'
'    <ここに処理を記述する>
'
'CleanExit:
'
'ErrorLog.TraceListPop
'Exit Sub 'or Function
'ErrorHandler:
'    ErrorLog.Raise Err.Number, Err.Description
'
'・子プロシージャの場合
'Const PROCEDURE_NAME As String = "＜プロシージャ名＞"
'Call ErrorLog.TraceListPush(MODULE_NAME, PROCEDURE_NAME)
'On Error GoTo ErrorHandler
'
'    <ここに処理を記述する>
'
'CleanExit:
'
'ErrorLog.TraceListPop
'Exit Sub 'or Function
'ErrorHandler:
'        Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
'
'その他 途中で行う処理で以下の物は左から右へ変換が必要
'On Error GoTo 0 → On Error GoTo ErrorHandler
'Exit FunctionまたはExit Sub → GoTo CleanExit 'Exit
'
'その他2 再帰関数には対応していない為、実装しない事。どうしてもやる場合コールスタックの閾値を外す
'----------------------------------------------------------------------------------------------------+

Option Explicit

Private Const ERR_MODULE_NAME As String = "ErrorLog"

'-----定義
'エラーファイル名
Private Const ERROR_LOGFILE_NAME As String = "errorlog.txt"
'スタックトレースの最大件数
Private Const MAX_STACK_COUNT As Long = 100
'コールスタックのオーバーフロー閾値
Private Const CALL_STACK_THRESHOLD As Long = 30
'ログファイルの最大MB数。超えた場合、oldファイルを残して新規ログファイルを作成する
Private Const MaxSizeMB As Long = 3

'-----モジュール変数
Private currentModuleName_ As String
Private currentProcedureName_ As String
Private callStack As Collection
Private stackTrace As Collection
Private otherInfo_ As String
Private removeCount As Long

Private Enum stackList
    mCallStack
    mStackTrace
End Enum

'予備用の変数。ダイナミックにデータを送りたい場合など
Public Property Let OtherInfo(ByVal OtherInfo As String)
    otherInfo_ = OtherInfo
End Property


'----------------------------------------------------------------------------------------------------+
'----------Title：Raise
'----------概要：エラーログを出力する
'-----Arg1：エラーオブジェクト
'-----Arg2：エラーメッセージ
'----------------------------------------------------------------------------------------------------+
Public Sub Raise(ByVal errorNumber As Long, ByVal errorMessage As String)

    Dim errorText As String
    Dim logPath As String
    Dim fileNumber As Long
    Dim strCallStack
    Dim strStackTrace
    
    'スタックトレースを文字列に書き起こす
    strCallStack = MakeStack(callStack, stackList.mCallStack)
    strStackTrace = MakeStack(stackTrace, stackList.mStackTrace)
    
    '古いスタックトレースの履歴を削除した場合、スタックトレースの末尾に件数を記述
    If removeCount > 0 Then
        strStackTrace = strStackTrace & vbCrLf & "              : " & vbTab & "*-容量オーバーの為、開始から" & removeCount & "件の履歴が削除されました。-*"
    End If
    
    'ログファイルの出力先ファイルパスを設定
    logPath = ThisWorkbook.path & "\" & ERROR_LOGFILE_NAME
    
    'ログファイルが重くなったら新規作成する
    Call RotateLogFile(logPath)
    
    'ログに書き込むテキストを設定
    errorText = "==================================================" & vbCrLf & _
                "  Timestamp   : " & Format(Now, "YYYY/MM/DD HH:mm:ss") & vbCrLf & _
                "  User        : " & Environ("USERNAME") & vbCrLf & _
                "  File        : " & ThisWorkbook.Name & vbCrLf & _
                "  Procedure   : " & currentModuleName_ & "." & currentProcedureName_ & vbCrLf & _
                "  Description : " & errorMessage & vbCrLf & _
                "  Info        : " & otherInfo_ & vbCrLf & _
                "  CallStack   : " & strCallStack & vbCrLf & _
                "  StackTrace  : " & strStackTrace & vbCrLf
    
    'ログファイルに出力
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
        Print #fileNumber, errorText
    Close #fileNumber
    
    '初期化
    Call ErrorLog.Initialize
    
    'エラーダイアログを発報
    AppActivate Application.Caption
    MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
            "エラー番号：" & errorNumber & vbCrLf & "概要：" & errorMessage, vbExclamation


End Sub

'-----初期化-----+
Sub Initialize()
    currentModuleName_ = ""
    currentProcedureName_ = ""
    Set callStack = Nothing
    Set stackTrace = Nothing
    OtherInfo = ""
    removeCount = 0
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TraceListPush
'----------概要：トレースリストに現在の処理状態をプッシュ
'-----Arg1  ：モジュール名
'-----Arg2  ：プロシージャ名
'----------------------------------------------------------------------------------------------------+
Public Sub TraceListPush(modName As String, procName As String)

    'コールスタックを追加
    If callStack Is Nothing Then Set callStack = New Collection
    callStack.Add modName & "." & procName
    'コールスタックのオーバーフローを確認
    Call TrimStack(callStack, CALL_STACK_THRESHOLD, True)
    
    'スタックトレースに履歴を残し、最大件数を確認
    If stackTrace Is Nothing Then Set stackTrace = New Collection
    stackTrace.Add "[+] " & modName & "." & procName
    Call TrimStack(stackTrace, MAX_STACK_COUNT)

    '現在実行中のモジュール.プロシージャを設定
    currentModuleName_ = modName
    currentProcedureName_ = procName

End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TraceListPop
'----------概要：トレースリストから現在の処理状態を削除する
'-----Arg：なし
'----------------------------------------------------------------------------------------------------+
Public Sub TraceListPop()
    Dim returnParent As Variant
    
    'コールスタックを削除
    If callStack Is Nothing Then Exit Sub
    callStack.Remove callStack.count
    
    'コールスタックが子から親に移る場合
    If callStack.count > 0 Then
        'スタックトレースを親に戻す
        If stackTrace Is Nothing Then Set stackTrace = New Collection
        stackTrace.Add "[-] " & callStack(callStack.count)
        Call TrimStack(stackTrace, MAX_STACK_COUNT)
        
        'コールスタックから親のモジュール.プロシージャ名を取得
        returnParent = callStack(callStack.count)
        
        '現在実行中のモジュール.プロシージャを親に戻す
        returnParent = Split(returnParent, ".")
        currentModuleName_ = returnParent(0)
        currentProcedureName_ = returnParent(1)

    End If

End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：MakeStack
'----------概要：トレースリストを文字列に書き起こして返却する
'-----Arg：トレースリストのコレクション
'----------------------------------------------------------------------------------------------------+
Private Function MakeStack(ByVal stackList As Collection, ByVal stackType As stackList) As String
    Dim i As Long
    Dim tmpStack As String

    If stackList Is Nothing Or stackList.count = 0 Then
        tmpStack = ""
    Else
    
        If stackType = mCallStack Then
            tmpStack = stackList(1)
            For i = 2 To stackList.count
                tmpStack = tmpStack & " -> " & stackList(i)
            Next i
        Else
            tmpStack = 1 & vbTab & stackList(1)
            For i = 2 To stackList.count
                tmpStack = tmpStack & vbCrLf & "              : " & i & vbTab & stackList(i)
            Next i
        End If
    End If
    
    MakeStack = tmpStack
End Function


'----------Title：RotateLogFile
'----------概要：ログファイルが引数のサイズを超えていたらリネームして退避する
'-----Arg：ログファイルパス
'----------------------------------------------------------------------------------------------------+
Private Sub RotateLogFile(ByVal logPath As String)

    Dim oldLogPath As String: oldLogPath = Replace(logPath, ".txt", "_old.txt")
    Dim maxSizeBytes As Long: maxSizeBytes = MaxSizeMB * 1024 * 1024&
    Dim fileNumber As Long
    On Error Resume Next 'ファイル不在時のDir/FileLenエラー回避
    
    'ファイルが存在し、かつサイズオーバーしている場合
    If Dir(logPath) <> "" Then
        If FileLen(logPath) > maxSizeBytes Then
            '1世代前のバックアップがあれば削除
            If Dir(oldLogPath) <> "" Then Kill oldLogPath
            '現在のファイルをバックアップ名に変更
            Name logPath As oldLogPath
            
            '新しいログファイルの先頭にオールドファイル生成の旨を書き込む
            fileNumber = FreeFile
            Open logPath For Append As #fileNumber
                Print #fileNumber, "--- [System Note: Log rotated at " & Now & " due to size limit (3MB)] ---"
            Close #fileNumber
            
            Debug.Print "Log rotated: " & oldLogPath & " created."
        End If
    End If
    
    On Error GoTo 0
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TrimStack
'----------概要：スタックトレースの件数を監視し、最大数を超えた場合、古いスタックトレースを削除する
'-----Arg：スタックトレースのコレクション
'----------------------------------------------------------------------------------------------------+
Private Sub TrimStack(ByRef stackList As Collection, ByVal stackThreshold As Long, Optional ByVal raiseFlg As Boolean = False)
    If stackList Is Nothing Then Exit Sub

    Do While stackList.count > stackThreshold
        If raiseFlg Then Err.Raise Number:=513, Description:="スタックリストが最大数(" & stackThreshold & ")を超過しています"
        stackList.Remove 1
        removeCount = removeCount + 1
    Loop
    
End Sub
