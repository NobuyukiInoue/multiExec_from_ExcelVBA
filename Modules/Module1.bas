Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------------------------------------
'Win32 API �֐��̐錾
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
'Win32 API �萔�̐錾
'------------------------------------------------------------------------
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const INFINITE As Long = &HFFFF


'------------------------------------------------------------------------
' �R�}���h������s���C��(Shell�֐���)
'------------------------------------------------------------------------
Public Sub multiExec_pidSearch_Shell(T_WS As Worksheet, processesMAX As Long, retryCountMAX As Long, getInterval As Long, waitEXIT As Boolean)
    '----------------------------------
    ' �ΏۃR�}���h��̓ǂݍ���
    '----------------------------------
    T_WS.Activate
    
    Dim T_WR As Range
    Set T_WR = T_WS.Range("A2")
    
    Dim CmdStr() As String
    CmdStr = set_CmdsList(T_WR)
    
    '----------------------------------
    ' �ΏۃR�}���h�̎��s
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
        ' �ΏۃR�}���h�̋N���v���Z�X��������������܂ő҂�
        '------------------------------------------------------------
        wait_enableExec processesMAX, retryCountMAX, getInterval, i, pid_list
       
        '------------------------------------------------------------
        ' �󂫂̃v���Z�X�ԍ��ۑ�����擾
        '------------------------------------------------------------
        Dim p As Long
        For p = LBound(pid_list) To UBound(pid_list)
           
           If (pid_list(p) = 0) And (i <= MAX_i) Then
                If CmdStr(i) <> "" Then
                    '------------------------------------------------------------
                    ' ���s�ΏۃR�}���h�̃Z���Ɉړ����A�X�e�[�^�X�o�[�ɏڍׂ�\��
                    '------------------------------------------------------------
                    T_WR.Cells(i + 1, 1).Activate
                    Application.StatusBar = Cells(T_WR.Row + i, 1).Address & " �̃R�}���h��������... [" & CmdStr(i) & "]"
                            
                    '------------------------------------------------------------
                    ' �R�}���h���s
                    '------------------------------------------------------------
                    On Error Resume Next
                    pid_list(p) = Shell(CmdStr(i), vbNormalFocus)
                    
                    If waitEXIT = True Then
                        '--------------------------------------------------------
                        ' �R�}���h���I���܂őҋ@����
                        '--------------------------------------------------------
                        If pid_list(p) <> 0 Then
                            ' �v���Z�X�̃V�O�i���҂�
                            Call WaitForSingleObject(pid_list(p), INFINITE)
                            
                            ' �v���Z�X�N���[�Y
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
    ' �I������
    '------------------------------------------------------------
    Range("A1").Activate
    Worksheets(1).Activate
    Application.StatusBar = False
    
    MsgBox "�������I���܂����B", vbOKOnly, "�R�}���h����s�I��"
    
End Sub


'------------------------------------------------------------------------
' �R�}���h������s���C��(WSH��)
'
' WSH�̎Q�Ɛݒ�)
' �P�D�u�c�[���v-->�u�Q�Ɛݒ�v���J��
' �Q�D�uWindows Script Host Object Model�v��I��
'------------------------------------------------------------------------
Public Sub multiExec_pidSearch_WSH(T_WS As Worksheet, processesMAX As Long, retryCountMAX As Long, getInterval As Long, waitEXIT As Boolean, WSHmethod As String)
    Dim sh As New IWshRuntimeLibrary.WshShell  '// WshShell�N���X�I�u�W�F�N�g
    Dim ex As WshExec                          '// Exec���\�b�h�߂�l
    
    '----------------------------------
    ' �ΏۃR�}���h��̓ǂݍ���
    '----------------------------------
    T_WS.Activate
    
    Dim T_WR As Range
    Set T_WR = T_WS.Range("A2")
    
    Dim CmdStr() As String
    CmdStr = set_CmdsList(T_WR)
    
    '----------------------------------
    ' �ΏۃR�}���h�̎��s
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
        ' �ΏۃR�}���h�̋N���v���Z�X��������������܂ő҂�
        '------------------------------------------------------------
        wait_enableExec processesMAX, retryCountMAX, getInterval, i, pid_list
       
        '------------------------------------------------------------
        ' �󂫂̃v���Z�X�ԍ��ۑ�����擾
        '------------------------------------------------------------
        Dim p As Long
        For p = LBound(pid_list) To UBound(pid_list)
           
           If (pid_list(p) = 0) And (i <= MAX_i) Then
                If CmdStr(i) <> "" Then
                    '------------------------------------------------------------
                    ' ���s�ΏۃR�}���h�̃Z���Ɉړ����A�X�e�[�^�X�o�[�ɏڍׂ�\��
                    '------------------------------------------------------------
                    T_WR.Cells(i + 1, 1).Activate
                    Application.StatusBar = Cells(T_WR.Row + i, 1).Address & " �̃R�}���h��������... [" & CmdStr(i) & "]"
                            
                    '------------------------------------------------------------
                    ' �R�}���h���s
                    '------------------------------------------------------------
                    On Error Resume Next
                    
                    If WSHmethod = "WSH(Exec)" Then
                    
                        '---------------------------------
                        ' Exec���\�b�h�i���_�C���N�g�s�j
                        '---------------------------------
                        Set ex = sh.Exec(CmdStr(i))
                    
                        If ex <> Null Then
                            pid_list(p) = ex.ProcessID
                            
                            If waitEXIT = True Then
                                '--------------------------------------------------------
                                ' �R�}���h���I���܂őҋ@����
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
                        ' Run���\�b�h�i���_�C���N�g�j
                        ' �X�e�[�^�X�̎擾�͑��̕��@�ŁB
                        '---------------------------------
                        If waitEXIT = True Then
                            sh.Run CmdStr(i), 1, True
                        Else
                            sh.Run CmdStr(i), 1, False
                        End If
                    
                    Else
                    
                        MsgBox "WSH�̃��\�b�h�̎w�肪��`�O�ł��B", vbOKOnly, "WSH���s�G���["
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
    ' �I������
    '------------------------------------------------------------
    Range("A1").Activate
    Worksheets(1).Activate
    Application.StatusBar = False
    
    MsgBox "�������I���܂����B", vbOKOnly, "�R�}���h����s�I��"
    
End Sub


'------------------------------------------------------------------------
' �w��V�[�g��A��ɃZ�b�g����Ă���R�}���h���z��ɓǂݍ���
'------------------------------------------------------------------------
Private Function set_CmdsList(WR As Range) As String()
    Dim CmdStr() As String
        
    Dim i As Long
    i = 0
    
    Do While WR.Cells(i + 1, 1).Value <> ""
            
        ReDim Preserve CmdStr(i)
        
        If Left(WR.Cells(i + 1, 1).Value, 1) = "#" Then
            
            '------------------------------------------------
            ' "#"����n�܂�Z���̓R�����g�Z���Ƃ��Ĉ������߁A
            ' �R�}���h�̓Z�b�g���Ȃ�
            '------------------------------------------------
            CmdStr(i) = ""
        
        ElseIf Left(WR.Cells(i + 1, 1).Value, 2) = ".\" Then
            
            '------------------------------------------------
            ' ���΃p�X�̏ꍇ�́A���̃u�b�N�̃p�X���w�肷��
            '------------------------------------------------
            CmdStr(i) = """" & ThisWorkbook.Path & "\" & Replace(WR.Cells(i + 1, 1).Value, ".\", "") & """"
        
        Else
            
            '------------------------------------------------
            ' �ΏۃZ���̃R�}���h��z��Ɋi�[����
            '------------------------------------------------
            CmdStr(i) = WR.Cells(i + 1, 1).Value
        
        End If
        
        i = i + 1
    Loop

    '--------------------------------------------
    ' �R�}���h����Z�b�g�����z���Ԃ�
    '--------------------------------------------
    set_CmdsList = CmdStr
    
End Function


'------------------------------------------------------------------------
' �ΏۃR�}���h�̋N���v���Z�X��������������܂ő҂�
'------------------------------------------------------------------------
Private Sub wait_enableExec(�}�N�������N��������� As Long, PID�擾���g���C�� As Long, �v���Z�X�ꗗ�擾�Ԋu As Long, i_Row As Long, ByRef pid_list() As Integer)
    
    Dim loopCount As Long
    loopCount = 0
               
    Do While True
        
        '----------------------------------------------------------------
        ' �v���Z�X�ꗗ����Ώۃv���Z�X��PID����������
        '----------------------------------------------------------------
        Dim pCount As Long
        pCount = 0
        
        Dim i As Long
        For i = LBound(pid_list) To UBound(pid_list)
            
            If IsExist_targetProcesses(pid_list(i)) Then
            
                '------------------------------------------
                ' ���������ꍇ�̓J�E���g����
                '------------------------------------------
                pCount = pCount + 1
            
            Else
                '------------------------------------------
                ' ������Ȃ������ꍇ�͏I�������Ɣ��f���A
                ' PID������������
                '------------------------------------------
                pid_list(i) = 0
            
            End If
        
        Next i
        
        If pCount < �}�N�������N��������� Then
            
            '------------------------------------------------------------
            ' �}�N�������N���������������Ă���ꍇ�́A���̏�����
            '------------------------------------------------------------
            Exit Sub
        
        End If
        
        loopCount = loopCount + 1
                        
        If loopCount >= PID�擾���g���C�� Then
            
            '------------------------------------------------------------
            ' ��莞�ԑ҂��Ă��v���Z�X���̏���������Ȃ������ꍇ
            '------------------------------------------------------------
            Dim msgRet As Long
            msgRet = MsgBox("��莞�ԑ҂��Ă��v���Z�X���̏���������܂���ł����B" & vbCrLf _
                                       & "targetCmd�̏I�����܂��҂��܂����H" & vbCrLf _
                                       & vbCrLf _
                                       & "���R�[�h�ԍ� : " & i_Row & vbCrLf _
                                       & "���݂̑ΏۃR�}���h��̋N���v���Z�X�� : " & pCount, _
                                       vbYesNo, "TeraTerm�}�N���N���v���Z�X���󂫑҂����Ԓ���")
            
            If msgRet <> vbYes Then
                
                '----------------------
                ' �}�N�����I������
                '----------------------
                End
            
            End If
            
            loopCount = 0
        
        Else
            
            '-------------------------------------------------------------
            ' [�v���Z�X�ꗗ�擾�Ԋu]�b�Ԓ�~����B
            '-------------------------------------------------------------
            DoEvents
            Application.Wait Now() + TimeSerial(0, 0, �v���Z�X�ꗗ�擾�Ԋu)
        
        End If
    Loop

End Sub


'------------------------------------------------------------------------
' �֐� : �w��PID�̃v���Z�X�����݂��Ă��邩���ׂ�
' https://msdn.microsoft.com/en-us/library/aa394372(v=vs.85).aspx
'------------------------------------------------------------------------
Public Function IsExist_targetProcesses(target_pid As Integer) As Boolean
    
    '--------------------------------------------
    ' target_pid�����w��(0)�̏ꍇ
    '--------------------------------------------
    If target_pid = 0 Then
        
        IsExist_targetProcesses = False
        Exit Function
    
    End If
    
    '--------------------------------------------
    ' WMI Win32_Process class�̃I�u�W�F�N�g����
    ' PID����������
    '--------------------------------------------
    Dim Locator: Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Dim Server: Set Server = Locator.ConnectServer
    Dim objSet: Set objSet = Server.ExecQuery("Select * From Win32_Process")
    Dim obj

    For Each obj In objSet
        If obj.ProcessID = target_pid Then
            
            '------------------------------------
            ' ���������ꍇ
            '------------------------------------
            IsExist_targetProcesses = True
            Exit Function
        
        End If
    Next
    
    IsExist_targetProcesses = False

End Function
