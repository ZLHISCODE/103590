Attribute VB_Name = "mdlMain"
Option Explicit
'==================================================================================================
'��д           lshuo
'����           2018/12/25
'ģ��           mdlMain
'˵��
'==================================================================================================
Private Const mstrCurModule     As String = "mdlMain"           '��ǰģ������

Public gobjFSO                  As New FileSystemObject                 'ȫ�ֵ��ļ��������
Public gobjRegister             As New clsRegister
Public gstrCommand              As String                           '����������
Public gblnSilence              As Boolean                          '�Ƿ�Ĭģʽ����
Public gstrServer               As String                           'Ҫִ�еķ�����
Public gblnInIDE                As Boolean                          '�Ƿ�Դ�뻷��
Public Const gstrSysName        As String = "�������"
Public gblnAsk                  As Boolean
Public glngSec                  As Long
Public glngLastTick             As Long

Public gblnShow                 As Boolean

Sub Main()
    Dim i           As Long
    Dim arrTmp      As Variant
    gblnInIDE = IsDesinMode
    AnalyzeCommandlineParameters
    gblnAsk = gstrServer = "*"
    If gstrServer <> "" Then
        gstrServer = GetServer(gstrServer)
    End If
    If gstrServer = "" Then End
    If Not gblnSilence And gblnAsk Then
        If frmMsgBox.ShowMsgBox(gstrSysName, "��ǰ�ͻ��˴������������δִ�е��ӳٽű����Ƿ�����ִ�У�", "!��(&Y),��(&N)", vbQuestion) = "��" Then
            End
        End If
    End If
    glngSec = 50
    gblnShow = True
    glngLastTick = GetTickCount
    arrTmp = Split(gstrServer, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i = LBound(arrTmp) Then
            Do While glngSec > 0
                Call ShowFlash("���ڼ�������" & arrTmp(i), , , arrTmp(i) & "")
                DoEvents
                Call Sleep(100)
                glngSec = glngSec - 1
            Loop
            gblnShow = False
        Else
            Call ShowFlash("���ڼ�������" & arrTmp(i), , , arrTmp(i) & "")
        End If
        Server = arrTmp(i)
        Do While True
            '��ִֹ�й����������˽ű�
            If RunUpgradeAfter Then
                Exit Do
            End If
        Loop
    Next
    Unload frmFlash
End Sub

'����������      ����
'-RunSVR         ִ�еķ���������Ϊ*�����Զ��������еķ�����������������Զ��ŷָ�
'-SILENCE        T-��Ĭ��ʽ
'ʾ����-RunAfter=ORCL -SILENCE=T
Public Sub AnalyzeCommandlineParameters(Optional ByVal strParams As String)
    Dim cSwitch As String, Path As String

    If IsMissing(strParams) = False Then
        CommandLine = strParams & " " & VBA.Command$
    Else
        CommandLine = VBA.Command$
    End If
    If Len(CommandLine) = 0 Then
        CommandLine = "-RUNSVR=* -SILENCE=F"
    End If
    gblnSilence = UCase$(CommandSwitch("SILENCE", False)) = "T"
    gstrServer = UCase$(CommandSwitch("RUNSVR", False))
End Sub

'--------------------------------------------------------------------------------------------------
'����           GetServer
'����           �жϲ���ȡ�����ӳٽű��ķ�����
'����ֵ         String
'����б�:
'������         ����                    ˵��
'
'-------------------------------------------------------------------------------------------------
Public Function GetServer(ByVal strServer As String) As String
    Dim strTmpServer        As String, strTmp   As String
    Dim objFile             As File
    Dim arrTmp              As Variant
    Dim i                   As Long
    
    If strServer = "*" Then
        If gobjFSO.FolderExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile") Then
            For Each objFile In gobjFSO.GetFolder(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile").Files
                If UCase(objFile.Name) Like "RUNAFTER_*.SQL" Then
                    strTmp = Mid$(objFile.Name, Len("RUNAFTER_*"))
                    strTmp = Trim(Mid$(strTmp, 1, Len(strTmp) - 4))
                    If strTmp <> "" Then
                        arrTmp = Split(strTmp, "_")
                        If Not IsDate(FullDate(arrTmp(UBound(arrTmp)))) Then
                            strTmpServer = strTmpServer & "," & strTmp
                        End If
                    End If
                End If
            Next
        End If
    Else
        arrTmp = Split(strServer, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If gobjFSO.FileExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & arrTmp(i) & ".SQL") Then
                strTmpServer = strTmpServer & "," & arrTmp(i)
            End If
        Next
    End If
    If strTmpServer <> "" Then strTmpServer = Mid$(strTmpServer, 2)
    GetServer = strTmpServer
End Function
