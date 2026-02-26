'----------------------------------------------------------------------------------------------------+
'-------------------------------------------*使用方法*-----------------------------------------------+
'----------------------------------------------------------------------------------------------------+
'各モジュールのトップステートメントに以下の定数を記述
'   Private Const MODULE_NAME As String = "＜モジュール名＞"
'
'各プロシージャの最初に以下の定数､トラップを記述
'   Const PROCEDURE_NAME As String = "＜プロシージャ名＞"
'   Call TraceListPush(MODULE_NAME, PROCEDURE_NAME)
'   On Error GoTo ErrorHandler
'
'各プロシージャの末尾に以下のスクリプトを記述
'CleanExit:
'    ErrorLog.TraceListPop
'    Exit Sub
'
'ErrorHandler:
'    If Err.Number <> 0 Then
'        AppActivate Application.Caption
'        MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
'                "エラー番号：" & Err.Number & vbCrLf & "概要：" & Err.Description, vbExclamation
'        ErrorLog.Raise Err, Err.Description
'        ErrorLog.TraceListPop
'        'デバッグモード。以後エラー箇所をループする
'        If ErrorLog.DEBUG_MODE Then Stop: Resume CleanExit Else End
'    End If
'
'他 途中で行う処理で以下の物は左から右へ変換が必要
'Exit FunctionまたはExit Sub → GoTo CleanExit 'Exit
'On Error GoTo 0 → On Error GoTo ErrorHandler
'End → ErrorLog.TraceListPop:End 'Exit
'----------------------------------------------------------------------------------------------------+


Option Explicit

Private Const ERR_MODULE_NAME As String = "ErrorLog"
Private Const ERROR_LOGFILE_NAME As String = "errorlog.txt"
Private currentModuleName_ As String
Private currentProcedureName_ As String
' Collectionに変更
Private stackList As Collection
Private callStack_ As String
Private stackTrace_ As String

'デバッグモード
Public Const DEBUG_MODE As Boolean = False


'----------------------------------------------------------------------------------------------------+
'----------Title：Raise
'----------概要：エラーログを出力する
'-----Arg1：エラーオブジェクト
'-----Arg2：エラーメッセージ
'----------------------------------------------------------------------------------------------------+
Public Sub Raise(ByVal errObj As ErrObject, ErrorMessage As String)
    Dim errorText As String
    Dim logPath As String
    Dim fileNumber As Integer
    On Error Resume Next
    
    logPath = ThisWorkbook.Path & "\" & ERROR_LOGFILE_NAME
    
    errorText = "==================================================" & vbCrLf & _
                "  Timestamp   : " & Format(Now, "YYYY/MM/DD HH:mm:ss") & vbCrLf & _
                "  User         : " & Environ("USERNAME") & vbCrLf & _
                "  File         : " & ThisWorkbook.Name & vbCrLf & _
                "  Procedure    : " & currentModuleName_ & "." & currentProcedureName_ & vbCrLf & _
                "  Description : " & ErrorMessage & vbCrLf & _
                "  CallStack    : " & callStack_ & vbCrLf & _
                "  StackTrace   : " & stackTrace_ & vbCrLf
    
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
        Print #fileNumber, errorText
    Close #fileNumber
    
    '初期化
    currentModuleName_ = ""
    currentProcedureName_ = ""
    callStack_ = ""
    stackTrace_ = ""
    Set stackList = Nothing
    
    If Err.Number <> 0 Then Debug.Print (ERR_MODULE_NAME & ".Raise " & Err.Description)
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TraceListPush
'----------概要：トレースリストに現在の処理状態をプッシュ
'-----Arg1  ：モジュール名
'-----Arg2  ：プロシージャ名
'----------------------------------------------------------------------------------------------------+
Public Sub TraceListPush(modName As String, procName As String)
    ' VBA標準の Collection を生成
    If stackList Is Nothing Then Set stackList = New Collection
    
    currentModuleName_ = modName
    currentProcedureName_ = procName
    
    ' 末尾に追加
    stackList.Add modName & "." & procName
    
    ' コールスタックを更新
    UpdateCallStack
    
    ' スタックトレースを更新
    If stackTrace_ = "" Then
        stackTrace_ = modName & "." & procName
    Else
        stackTrace_ = stackTrace_ & " -> " & modName & "." & procName
    End If
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TraceListPop
'----------概要：トレースリストから現在の処理状態を削除する
'-----Arg：なし
'----------------------------------------------------------------------------------------------------+
Public Sub TraceListPop()
    Dim returnParent As Variant
    If stackList Is Nothing Then Exit Sub
    
    'コールスタックが子から親に移る場合
    If stackList.count > 0 Then
        stackList.Remove stackList.count

        'コールスタックが子から親に移る場合
        If stackList.count > 0 Then
            'スタックトレースを親に戻す
            returnParent = stackList(stackList.count)
            stackTrace_ = stackTrace_ & " -> " & returnParent
            
            '現在実行中のプロシージャを親に戻す
            returnParent = Split(returnParent, ".")
            currentModuleName_ = returnParent(0)
            currentProcedureName_ = returnParent(1)
        End If
    End If
    UpdateCallStack
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：UpdateCallStack
'----------概要：トレースリストを文字列に書き起こすサブルーチン
'-----Arg：なし
'----------------------------------------------------------------------------------------------------+
Private Sub UpdateCallStack()
    If stackList Is Nothing Or stackList.Count = 0 Then
        callStack_ = ""
    Else
        Dim i As Long
        Dim tmpStack As String
        tmpStack = ""
        For i = 1 To stackList.Count
            tmpStack = tmpStack & IIf(i = 1, "", " -> ") & stackList(i)
        Next i
        callStack_ = tmpStack
    End If
End Sub
