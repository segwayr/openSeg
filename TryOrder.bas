Option Explicit

Private Const MODULE_NAME As String = "TryOrder"
Private webDriver_ As Selenium.WebDriver
Private target_ As String
Private inputText_ As String
Private selector_ As String
Private indexNumber_ As Long
Private timeoutSec_ As Double
Private raiseFlag_ As Boolean

'Execute、Valueメソッドで実行したいメソッド名
Public Enum methodType
    mClick
    mClear
    mSendKeys
    mGetValue
    mIsSelected
    mInnerText
    mExecuteScript
End Enum

'+----------Start アクセサ----------+

'DOMを再取得する為のアクセサ
Private Property Get webElm() As Object

Select Case Selector
    Case "ID": Set webElm = WebDriver.FindElementsById(Target)
    Case "Name": Set webElm = WebDriver.FindElementsByName(Target)
    Case "Css": Set webElm = WebDriver.FindElementsByCss(Target)
    Case "XPath": Set webElm = WebDriver.FindElementsByXPath(Target)
    Case "LinkText": Set webElm = WebDriver.FindElementsByLinkText(Target)
End Select

End Property

'findElementsBy...で取得する対象のセレクタ名を格納した配列。読み取り専用プロパティ
Private Property Get SELECTOR_LIST() As Variant
    SELECTOR_LIST = Array("ID", "Name", "Css", "XPath", "LinkText")
End Property



'-----メンバ変数
'WebDriver
Public Property Set WebDriver(ByRef WebDriver As Selenium.WebDriver)
    If WebDriver Is Nothing Then
        Err.Raise 513, Description:="WebDriverがセットされていません"
    Else
        Set webDriver_ = WebDriver
    End If
End Property

Private Property Get WebDriver() As Selenium.WebDriver
    If webDriver_ Is Nothing Then
        Err.Raise 513, Description:="WebDriverがセットされていません"
    Else
        Set WebDriver = webDriver_
    End If
End Property


'属性
Public Property Let Target(ByVal Target As String)
    target_ = Target
End Property

Private Property Get Target() As String
    If target_ = "" Then
        Err.Raise 513, Description:="属性がセットされていません"
    Else
        Target = target_
    End If
End Property

'文字列。スクリプトや入力する値
Public Property Let InputText(ByVal InputText As String)
    inputText_ = InputText
End Property

Private Property Get InputText() As String
        InputText = inputText_
End Property

'セレクタ。ByNameやByCss等
Public Property Let Selector(ByVal Selector As String)
    If IsError(Application.Match(Selector, SELECTOR_LIST, 0)) Then
        Err.Raise 513, Description:="不正なセレクターです: " & Selector
    End If
    selector_ = Selector
End Property

Private Property Get Selector() As String
    If selector_ = "" Then
        Selector = "Name"
    Else
        Selector = selector_
    End If
End Property


'インデックス番号。Item(n)
Public Property Let IndexNumber(ByVal IndexNumber As Long)
    indexNumber_ = IndexNumber
End Property

Private Property Get IndexNumber() As Long
    IndexNumber = indexNumber_
End Property

'文字列。スクリプトや入力する値
Public Property Let TimeoutSec(ByVal TimeoutSec As Double)
    timeoutSec_ = TimeoutSec
End Property

Private Property Get TimeoutSec() As Double
    If TimeoutSec < 0 Then
        TimeoutSec = 10
    Else
        TimeoutSec = timeoutSec_
    End If
End Property

'文字列。スクリプトや入力する値
Public Property Let RaiseFlag(ByVal RaiseFlag As Boolean)
    raiseFlag_ = RaiseFlag
End Property
Private Property Get RaiseFlag() As Boolean
    RaiseFlag = raiseFlag_
End Property


'+----------End アクセサ----------+

'コンストラクター
Sub Class_Initialize()
    Const PROCEDURE_NAME As String = "Class_Initialize"
    Call TraceListPush(MODULE_NAME, PROCEDURE_NAME)
    On Error GoTo ErrorHandler
    
    Call ClearProperties
    
CleanExit:
    ErrorLog.TraceListPop
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
                "エラー番号：" & Err.Number & vbCrLf & "概要：" & Err.Description, vbExclamation
        ErrorLog.Raise Err, Err.Description
        ErrorLog.TraceListPop
        'デバッグモード。以後エラー箇所をループする
        If ErrorLog.DEBUG_MODE Then Stop: Resume CleanExit Else End
    End If
End Sub


'-----以下メソッド


'----------------------------------------------------------------------------------------------------+
'----------Title：ClearProperties
'----------概要：プロパティを初期化する
'-----Arg：なし
'-----戻り値：なし
'----------------------------------------------------------------------------------------------------+
Public Sub ClearProperties()
    Set webDriver_ = Nothing
    target_ = ""
    inputText_ = ""
    indexNumber_ = 1
    selector_ = "Name"
    timeoutSec_ = 10
    raiseFlag_ = True
End Sub


'----------------------------------------------------------------------------------------------------+
'----------Title：DoAction
'----------概要：第一引数で指定したメソッドを成功するまで実行し続ける（Enum参照）
'-----Arg1：実行するメソッド
'-----Arg2：検索のタイムアウト秒数時間。デフォルトで10秒間
'-----戻り値：タイムアウトまでに実行できたかどうかの真偽値
'----------------------------------------------------------------------------------------------------+
Public Function DoAction(ByVal method As methodType) As Boolean
    Const PROCEDURE_NAME As String = "DoAction"
    Call TraceListPush(MODULE_NAME, PROCEDURE_NAME)
    On Error GoTo ErrorHandler
    Dim startTime As Double
    
    startTime = Timer
    Do While Timer - startTime < TimeoutSec
        On Error Resume Next
        
        'プロパティに対して指定したメソッドを実行する
        Select Case method
            Case methodType.mExecuteScript: WebDriver.ExecuteScript InputText
            Case methodType.mClick:         webElm.Item(IndexNumber).Click
            Case methodType.mClear:         webElm.Item(IndexNumber).Clear
            Case methodType.mSendKeys:      webElm.Item(IndexNumber).SendKeys InputText
            Case Else: Err.Raise 513, Description:="不正なメソッドです。引数を確認してください。"
        End Select
        
        'メソッドが正常に実行された場合、Trueを返却して終了
        If Err.Number = 0 Then
            DoAction = True
            GoTo CleanExit 'Exit
        End If
        On Error GoTo ErrorHandler
        
        WebDriver.Wait 200
        DoEvents
    
    Loop
    
    DoAction = False
    If RaiseFlag Then Err.Raise Number:=513, Description:="ASTRA読込エラーです。通信状態を確認後、再度実行してください。"
    
CleanExit:
    ErrorLog.TraceListPop
    Exit Function

ErrorHandler:
    If Err.Number <> 0 Then
        AppActivate Application.Caption
        MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
                "エラー番号：" & Err.Number & vbCrLf & "概要：" & Err.Description, vbExclamation
        ErrorLog.Raise Err, Err.Description
        ErrorLog.TraceListPop
        'デバッグモード。以後エラー箇所をループする
        If ErrorLog.DEBUG_MODE Then Stop: Resume CleanExit Else End
    End If
    
End Function


'----------------------------------------------------------------------------------------------------+
'----------Title：GetResult
'----------概要：第一引数で指定したメソッドを成功するまで実行し続ける（Enum参照）
'-----Arg1：実行するメソッド
'-----戻り値：取得した値を返却
'----------------------------------------------------------------------------------------------------+
Public Function GetResult(ByVal method As methodType) As Variant
    Const PROCEDURE_NAME As String = "GetResult"
    Call TraceListPush(MODULE_NAME, PROCEDURE_NAME)
    On Error GoTo ErrorHandler
    
    Dim scriptStr As String
    Dim startTime As Double
    Dim textList As Collection
    startTime = Timer
    Do While Timer - startTime < TimeoutSec

        On Error Resume Next

        'プロパティに対して指定したメソッドを実行し、値を取得する
        Select Case method
            Case methodType.mExecuteScript
                scriptStr = InputText
                If Not LCase(scriptStr) Like "return *" Then
                    scriptStr = "return " & scriptStr
                End If
                GetResult = WebDriver.ExecuteScript(scriptStr)
                
            Case methodType.mGetValue:   GetResult = webElm.Item(IndexNumber).value
            Case methodType.mIsSelected: GetResult = webElm.Item(IndexNumber).IsSelected
            Case methodType.mInnerText: GetResult = GetText()
                    
            Case Else: Err.Raise 513, Description:="不正なメソッドです。引数を確認してください。"
        End Select

        'メソッドが正常に実行され、値が返却された場合、Trueを返却して終了
        If Err.Number = 0 Then
            GoTo CleanExit 'Exit
        End If
        On Error GoTo ErrorHandler
        
        WebDriver.Wait 200 ' 0.2秒刻みでチェック
        DoEvents
    
    Loop
    
    GetResult = False
    If RaiseFlag Then Err.Raise Number:=513, Description:="ASTRA読込エラーです。通信状態を確認後、再度実行してください。"
    
CleanExit:
    ErrorLog.TraceListPop
    Exit Function

ErrorHandler:
    If Err.Number <> 0 Then
        AppActivate Application.Caption
        MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
                "エラー番号：" & Err.Number & vbCrLf & "概要：" & Err.Description, vbExclamation
        ErrorLog.Raise Err, Err.Description
        ErrorLog.TraceListPop
        'デバッグモード。以後エラー箇所をループする
        If ErrorLog.DEBUG_MODE Then Stop: Resume CleanExit Else End
    End If
    
End Function


'以下、内部処理用メソッド
'----------------------------------------------------------------------------------------------------+
'----------Title：GetText
'----------概要：対象のオブジェクトのテキストを成功するまで取得し続ける
'-----Arg：なし
'-----戻り値：取得した値をコレクションで返却
'----------------------------------------------------------------------------------------------------+
Private Function GetText() As Variant
    Const PROCEDURE_NAME As String = "GetText"
    Call TraceListPush(MODULE_NAME, PROCEDURE_NAME)
    On Error GoTo ErrorHandler
    
    Dim results() As String
    Dim curObj As Object
    Dim startTime As Double
    Dim count As Long
    
    startTime = Timer
    Do While Timer - startTime < TimeoutSec
        On Error Resume Next
        count = 0
        Erase results
        
        ReDim results(0 To webElm.count - 1)
        For Each curObj In webElm
            results(count) = curObj.Text
            count = count + 1
        Next curObj
        
        
        '正しくテキストが抜けた場合
        If Err.Number = 0 Then
            GetText = results
            GoTo CleanExit 'Exit
        Else
            results = Empty
        End If
        On Error GoTo ErrorHandler
        
        WebDriver.Wait 200
        DoEvents
    
    Loop
    
    If RaiseFlag Then Err.Raise Number:=513, Description:="ASTRA読込エラーです。通信状態を確認後、再度実行してください。"
    
    
CleanExit:
    ErrorLog.TraceListPop
    Exit Function

ErrorHandler:
    If Err.Number <> 0 Then
        AppActivate Application.Caption
        MsgBox "エラーが発生したため中断しました。" & vbCrLf & _
                "エラー番号：" & Err.Number & vbCrLf & "概要：" & Err.Description, vbExclamation
        ErrorLog.Raise Err, Err.Description
        ErrorLog.TraceListPop
        'デバッグモード。以後エラー箇所をループする
        If ErrorLog.DEBUG_MODE Then Stop: Resume CleanExit Else End
    End If

End Function
