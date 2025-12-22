'----------------------------------------------------------------------------------------------------+
'-------------------------------------------*使用方法*-----------------------------------------------+
'----------------------------------------------------------------------------------------------------+'
''各モジュールのトップステートメントに以下の定数を記述
'     Private Const MODULE_NAME = "＜モジュール名＞"
'
''各プロシージャの最初に以下の定数､トラップを記述
'    Const PROCEDURE_NAME = "＜プロシージャ名＞"
'    Call TraceListPush(MODULE_NAME, PROCEDURE_NAME)
'    On Error GoTo ErrorHandler
''
''各プロシージャの末尾に以下のスクリプトを記述
'CleanExit:
'    ErrorLog.TraceListPop
'Exit Sub '※もしくはExit Function
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
'        "エラー番号：" & Err.Number & vbCrLf & "概要：" & Err.Description, vbExclamation
'        ErrorLog.Raise Err, Err.Description
'    End If
'    Resume CleanExit
'----------------------------------------------------------------------------------------------------+

Option Explicit

Private Const ERR_MODULE_NAME = "ErrorLog"
Private Const ERROR_LOGFILE_NAME = "errorlog.txt"

Private currentModuleName_ As String
Private currentProcedureName_ As String
Private stackList As Object
Private callStack_ As String
Private stackTrace_ As String


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
       
      'ログファイルの出力先ファイルパスを設定
      logPath = ThisWorkbook.Path & "\" & ERROR_LOGFILE_NAME
       
      'ログに書き込むテキストを設定
      errorText = "==================================================" & vbCrLf & _
                        "   Timestamp   : " & Format(Now, "YYYY/MM/DD HH:mm:ss") & vbCrLf & _
                        "   User        : " & Environ("USERNAME") & vbCrLf & _
                        "   File        : " & ThisWorkbook.Name & vbCrLf & _
                        "   Procedure   : " & currentModuleName_ & "." & currentProcedureName_ & vbCrLf & _
                        "   Description : " & ErrorMessage & vbCrLf & _
                        "   CallStack   : " & callStack_ & vbCrLf & _
                        "   StackTrace  : " & stackTrace_ & vbCrLf
       
      'ログファイルに出力
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
       
      'ErrorHandler
      If Err.Number <> 0 Then Debug.Print (ERR_MODULE_NAME & ".Raise " & Err.Description)

End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TraceListPush
'----------概要：トレースリストに現在の処理状態をプッシュ
'-----Arg1   ：モジュール名
'-----Arg2   ：プロシージャ名
'----------------------------------------------------------------------------------------------------+
Public Sub TraceListPush(modName As String, procName As String)
      If stackList Is Nothing Then Set stackList = CreateObject("System.Collections.ArrayList")
       
      currentModuleName_ = modName
      currentProcedureName_ = procName
       
      '「モジュール名.プロシージャ名」をスタックに追加
      stackList.Add modName & "." & procName
       
      'コールスタックを更新
      UpdateCallStack
       
      'スタックトレースを更新
      If stackTrace_ = "" Then
            stackTrace_ = modName & "." & procName
      Else
            stackTrace_ = stackTrace_ & " -> " & modName & "." & procName
      End If
       
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：TraceListPop
'----------概要：トレースリストから現在の処理状態を削除する
'-----Arg1：モジュール名
'-----Arg2：プロシージャ名
'----------------------------------------------------------------------------------------------------+
Public Sub TraceListPop()
      If stackList Is Nothing Then Exit Sub
       
      If stackList.Count > 0 Then
            stackList.RemoveAt stackList.Count - 1
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
            callStack_ = Join(stackList.ToArray(), " -> ")
      End If
End Sub

