Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------------------------------------
'Win32 API 関数の宣言
'------------------------------------------------------------------------
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

#End If

'------------------------------------------------------------------------
'Win32 API 定数の宣言
'------------------------------------------------------------------------
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const INFINITE As Long = &HFFFF


'------------------------------------------------------------------------
' コマンド並列実行メイン(Shell関数版)
'------------------------------------------------------------------------
Public Sub multiExec_pidSearch_Shell(T_WS As Worksheet, processesMAX As Long, retryCountMAX As Long, getInterval As Long, waitEXIT As Boolean)
    '----------------------------------
    ' 対象コマンド列の読み込み
    '----------------------------------
    T_WS.Activate
    
    Dim T_WR As Range
    Set T_WR = T_WS.Range("A2")
    
    Dim CmdStr() As String
    CmdStr = set_CmdsList(T_WR)
    
    '----------------------------------
    ' 対象コマンドの実行
    '----------------------------------
    Dim pid_list() As Integer
    ReDim Preserve pid_list(0 To processesMAX - 1)
    
    Dim MIN_i As Long
    Dim MAX_i As Long
    MIN_i = LBound(CmdStr)
    MAX_i = UBound(CmdStr)
    
    Dim i As Long
    i = MIN_i: Do While i <= MAX_i
    
        '------------------------------------------------------------
        ' 対象コマンドの起動プロセス数が上限を下回るまで待つ
        '------------------------------------------------------------
        wait_enableExec processesMAX, retryCountMAX, getInterval, i, pid_list
       
        '------------------------------------------------------------
        ' 空きのプロセス番号保存先を取得
        '------------------------------------------------------------
        Dim p As Long
        For p = LBound(pid_list) To UBound(pid_list)
           
           If (pid_list(p) = 0) And (i <= MAX_i) Then
                If CmdStr(i) <> "" Then
                    '------------------------------------------------------------
                    ' 実行対象コマンドのセルに移動し、ステータスバーに詳細を表示
                    '------------------------------------------------------------
                    T_WR.Cells(i + 1, 1).Activate
                    Application.StatusBar = Cells(T_WR.Row + i, 1).Address & " のコマンドを処理中... [" & CmdStr(i) & "]"
                            
                    '------------------------------------------------------------
                    ' コマンド実行
                    '------------------------------------------------------------
                    On Error Resume Next
                    pid_list(p) = Shell(CmdStr(i), vbNormalFocus)
                    
                    If waitEXIT = True Then
                        '--------------------------------------------------------
                        ' コマンドが終了まで待機する
                        '--------------------------------------------------------
                        If pid_list(p) <> 0 Then
                            ' プロセスのシグナル待ち
                            Call WaitForSingleObject(pid_list(p), INFINITE)
                            
                            ' プロセスクローズ
                            CloseHandle pid_list(p)
                        End If
                    
                    End If
                
                End If
                
                If i <= MAX_i Then
                    i = i + 1
                End If
           End If
        
        Next p
    Loop

    '------------------------------------------------------------
    ' 終了処理
    '------------------------------------------------------------
    Range("A1").Activate
    Worksheets(1).Activate
    Application.StatusBar = False
    
    MsgBox "処理が終わりました。", vbOKOnly, "コマンド列実行終了"
    
End Sub


'------------------------------------------------------------------------
' コマンド並列実行メイン(WSH版)
'
' WSHの参照設定)
' １．「ツール」-->「参照設定」を開く
' ２．「Windows Script Host Object Model」を選択
'------------------------------------------------------------------------
Public Sub multiExec_pidSearch_WSH(T_WS As Worksheet, processesMAX As Long, retryCountMAX As Long, getInterval As Long, waitEXIT As Boolean, WSHmethod As String)
    Dim sh As New IWshRuntimeLibrary.WshShell  '// WshShellクラスオブジェクト
    Dim ex As WshExec                          '// Execメソッド戻り値
    
    '----------------------------------
    ' 対象コマンド列の読み込み
    '----------------------------------
    T_WS.Activate
    
    Dim T_WR As Range
    Set T_WR = T_WS.Range("A2")
    
    Dim CmdStr() As String
    CmdStr = set_CmdsList(T_WR)
    
    '----------------------------------
    ' 対象コマンドの実行
    '----------------------------------
    Dim pid_list() As Integer
    ReDim Preserve pid_list(0 To processesMAX - 1)
    
    Dim MIN_i As Long
    Dim MAX_i As Long
    MIN_i = LBound(CmdStr)
    MAX_i = UBound(CmdStr)
    
    Dim i As Long
    i = MIN_i: Do While i <= MAX_i
    
        '------------------------------------------------------------
        ' 対象コマンドの起動プロセス数が上限を下回るまで待つ
        '------------------------------------------------------------
        wait_enableExec processesMAX, retryCountMAX, getInterval, i, pid_list
       
        '------------------------------------------------------------
        ' 空きのプロセス番号保存先を取得
        '------------------------------------------------------------
        Dim p As Long
        For p = LBound(pid_list) To UBound(pid_list)
           
           If (pid_list(p) = 0) And (i <= MAX_i) Then
                If CmdStr(i) <> "" Then
                    '------------------------------------------------------------
                    ' 実行対象コマンドのセルに移動し、ステータスバーに詳細を表示
                    '------------------------------------------------------------
                    T_WR.Cells(i + 1, 1).Activate
                    Application.StatusBar = Cells(T_WR.Row + i, 1).Address & " のコマンドを処理中... [" & CmdStr(i) & "]"
                            
                    '------------------------------------------------------------
                    ' コマンド実行
                    '------------------------------------------------------------
                    On Error Resume Next
                    
                    If WSHmethod = "WSH(Exec)" Then
                    
                        '---------------------------------
                        ' Execメソッド（リダイレクト不可）
                        '---------------------------------
                        Set ex = sh.Exec(CmdStr(i))
                    
                        If ex <> Null Then
                            pid_list(p) = ex.ProcessID
                            
                            If waitEXIT = True Then
                                '--------------------------------------------------------
                                ' コマンドが終了まで待機する
                                '--------------------------------------------------------
                                Do While (ex.Status = WshRunning)
                                    DoEvents
                                    Sleep 100
                                Loop
                            End If
                            
                            Set ex = Nothing
                        End If
                    
                    ElseIf WSHmethod = "WSH(Run)" Then
                    
                        '---------------------------------
                        ' Runメソッド（リダイレクト可）
                        ' ステータスの取得は他の方法で。
                        '---------------------------------
                        If waitEXIT = True Then
                            sh.Run CmdStr(i), 1, True
                        Else
                            sh.Run CmdStr(i), 1, False
                        End If
                    
                    Else
                    
                        MsgBox "WSHのメソッドの指定が定義外です。", vbOKOnly, "WSH実行エラー"
                        End
                        
                    End If
                    
                    
                
                End If
                
                If i <= MAX_i Then
                    i = i + 1
                End If
           End If
        
        Next p
    Loop

    '------------------------------------------------------------
    ' 終了処理
    '------------------------------------------------------------
    Range("A1").Activate
    Worksheets(1).Activate
    Application.StatusBar = False
    
    MsgBox "処理が終わりました。", vbOKOnly, "コマンド列実行終了"
    
End Sub


'------------------------------------------------------------------------
' 指定シートのA列にセットされているコマンド列を配列に読み込む
'------------------------------------------------------------------------
Private Function set_CmdsList(WR As Range) As String()
    Dim CmdStr() As String
        
    Dim i As Long
    i = 0
    
    Do While WR.Cells(i + 1, 1).Value <> ""
            
        ReDim Preserve CmdStr(i)
        
        If Left(WR.Cells(i + 1, 1).Value, 1) = "#" Then
            
            '------------------------------------------------
            ' "#"から始まるセルはコメントセルとして扱うため、
            ' コマンドはセットしない
            '------------------------------------------------
            CmdStr(i) = ""
        
        ElseIf Left(WR.Cells(i + 1, 1).Value, 2) = ".\" Then
            
            '------------------------------------------------
            ' 相対パスの場合は、このブックのパスを指定する
            '------------------------------------------------
            CmdStr(i) = """" & ThisWorkbook.Path & "\" & Replace(WR.Cells(i + 1, 1).Value, ".\", "") & """"
        
        Else
            
            '------------------------------------------------
            ' 対象セルのコマンドを配列に格納する
            '------------------------------------------------
            CmdStr(i) = WR.Cells(i + 1, 1).Value
        
        End If
        
        i = i + 1
    Loop

    '--------------------------------------------
    ' コマンド列をセットした配列を返す
    '--------------------------------------------
    set_CmdsList = CmdStr
    
End Function


'------------------------------------------------------------------------
' 対象コマンドの起動プロセス数が上限を下回るまで待つ
'------------------------------------------------------------------------
Private Sub wait_enableExec(マクロ同時起動数上限数 As Long, PID取得リトライ回数 As Long, プロセス一覧取得間隔 As Long, i_Row As Long, ByRef pid_list() As Integer)
    
    Dim loopCount As Long
    loopCount = 0
               
    Do While True
        
        '----------------------------------------------------------------
        ' プロセス一覧から対象プロセスのPIDを検索する
        '----------------------------------------------------------------
        Dim pCount As Long
        pCount = 0
        
        Dim i As Long
        For i = LBound(pid_list) To UBound(pid_list)
            
            If IsExist_targetProcesses(pid_list(i)) Then
            
                '------------------------------------------
                ' 見つかった場合はカウントする
                '------------------------------------------
                pCount = pCount + 1
            
            Else
                '------------------------------------------
                ' 見つからなかった場合は終了したと判断し、
                ' PIDを初期化する
                '------------------------------------------
                pid_list(i) = 0
            
            End If
        
        Next i
        
        If pCount < マクロ同時起動数上限数 Then
            
            '------------------------------------------------------------
            ' マクロ同時起動上限数を下回っている場合は、次の処理へ
            '------------------------------------------------------------
            Exit Sub
        
        End If
        
        loopCount = loopCount + 1
                        
        If loopCount >= PID取得リトライ回数 Then
            
            '------------------------------------------------------------
            ' 一定時間待ってもプロセス数の上限を下回らなかった場合
            '------------------------------------------------------------
            Dim msgRet As Long
            msgRet = MsgBox("一定時間待ってもプロセス数の上限を下回りませんでした。" & vbCrLf _
                                       & "targetCmdの終了をまだ待ちますか？" & vbCrLf _
                                       & vbCrLf _
                                       & "レコード番号 : " & i_Row & vbCrLf _
                                       & "現在の対象コマンド列の起動プロセス数 : " & pCount, _
                                       vbYesNo, "TeraTermマクロ起動プロセス数空き待ち時間超過")
            
            If msgRet <> vbYes Then
                
                '----------------------
                ' マクロを終了する
                '----------------------
                End
            
            End If
            
            loopCount = 0
        
        Else
            
            '-------------------------------------------------------------
            ' [プロセス一覧取得間隔]秒間停止する。
            '-------------------------------------------------------------
            DoEvents
            Application.Wait Now() + TimeSerial(0, 0, プロセス一覧取得間隔)
        
        End If
    Loop

End Sub


'------------------------------------------------------------------------
' 関数 : 指定PIDのプロセスが存在しているか調べる
' https://msdn.microsoft.com/en-us/library/aa394372(v=vs.85).aspx
'------------------------------------------------------------------------
Public Function IsExist_targetProcesses(target_pid As Integer) As Boolean
    
    '--------------------------------------------
    ' target_pidが未指定(0)の場合
    '--------------------------------------------
    If target_pid = 0 Then
        
        IsExist_targetProcesses = False
        Exit Function
    
    End If
    
    '--------------------------------------------
    ' WMI Win32_Process classのオブジェクトから
    ' PIDを検索する
    '--------------------------------------------
    Dim Locator: Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Dim Server: Set Server = Locator.ConnectServer
    Dim objSet: Set objSet = Server.ExecQuery("Select * From Win32_Process")
    Dim obj

    For Each obj In objSet
        If obj.ProcessID = target_pid Then
            
            '------------------------------------
            ' 見つかった場合
            '------------------------------------
            IsExist_targetProcesses = True
            Exit Function
        
        End If
    Next
    
    IsExist_targetProcesses = False

End Function
