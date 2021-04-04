Attribute VB_Name = "mdlMain"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2018/12/25
'模块           mdlMain
'说明
'==================================================================================================
Private Const mstrCurModule     As String = "mdlMain"           '当前模块名称

Public gobjFSO                  As New FileSystemObject                 '全局的文件处理对象
Public gobjRegister             As New clsRegister
Public gstrCommand              As String                           '命令行内容
Public gblnSilence              As Boolean                          '是否静默模式处理
Public gstrServer               As String                           '要执行的服务器
Public gblnInIDE                As Boolean                          '是否源码环境
Public Const gstrSysName        As String = "中联软件"
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
        If frmMsgBox.ShowMsgBox(gstrSysName, "当前客户端存在最近升级尚未执行的延迟脚本，是否现在执行？", "!是(&Y),否(&N)", vbQuestion) = "否" Then
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
                Call ShowFlash("正在检查服务器" & arrTmp(i), , , arrTmp(i) & "")
                DoEvents
                Call Sleep(100)
                glngSec = glngSec - 1
            Loop
            gblnShow = False
        Else
            Call ShowFlash("正在检查服务器" & arrTmp(i), , , arrTmp(i) & "")
        End If
        Server = arrTmp(i)
        Do While True
            '防止执行过程中新增了脚本
            If RunUpgradeAfter Then
                Exit Do
            End If
        Loop
    Next
    Unload frmFlash
End Sub

'命令行类型      处理
'-RunSVR         执行的服务器，若为*，则自动查找所有的服务器。多个服务器以逗号分隔
'-SILENCE        T-静默方式
'示例：-RunAfter=ORCL -SILENCE=T
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
'方法           GetServer
'功能           判断并获取存在延迟脚本的服务器
'返回值         String
'入参列表:
'参数名         类型                    说明
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
