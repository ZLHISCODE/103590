VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub InitClsPublic()
    
'    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
'    Set ParentWnd = objParent
    blnDataReceived = False
    
    
'    Set gcnOracle = New ADODB.Connection
'    With gcnOracle
'        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
'        .Open cnMain.ConnectionString
'    End With
'    Set gcnOracle = cnMain
    
'    With gcnAccess
'        If .State = adStateOpen Then .Close
'        .ConnectionString = "DBQ=" & App.Path & "\ZlLISLog.mdb;DefaultDir=" & App.Path & ";Driver={Microsoft Access Driver (*.mdb)}"
'        .Open
'    End With
'---  初始化公共部件


    '清空日志文件
    If gblnFromDB Then
        gblnClearData = gobjDatabase.GetPara("清空接收日志", glngSys, 1208, 1)
    Else
        gblnClearData = GetSetting("ZLSOFT", "公共模块\ZlLISSrv", "清空接收日志", 1)
    End If
    If gblnClearData Then
        On Error GoTo DBError
        If Dir(App.Path & "\zlLis*.LOG") <> "" Then
            Kill App.Path & "\zlLis*.LOG"
        End If
        
        On Error GoTo 0
    End If
    
    '初始酶标仪参数

    InitMBPara

    '-------------------------------------------------
    frmLISSrv.Show
    Exit Sub
DBError:
    Call WriteLog("clsPublic.InitClsPub", LOG_错误日志, Err.Number, Err.Description)
End Sub

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
'发送标本记录到仪器
'    frmLISSrv.Show
    SendSample = frmLISSrv.SendSample(lngDeviceID, strSampleDate, strSampleNO, strAdviceIDs, blnUndo, iType)
End Function

Public Sub GetResultFromFile(ByVal strFile As String, ByVal lngDeviceID As Long, ByVal strSampleNO As String, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))
'从文件获取检验结果
'   strFile：包含路径的文件名
'   lngDeviceID：检验设备ID
'   strSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（strSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。

    ResultFromFile strFile, lngDeviceID, strSampleNO, dtStart, dtEnd
End Sub

Private Sub Class_Terminate()
'    Unload frmLISSrv
'    With gcnOracle
'        If .State = adStateOpen Then .Close
'    End With
'    With gcnAccess
'        If .State = adStateOpen Then .Close
'    End With
'    Set gcnOracle = Nothing: Set gcnAccess = Nothing
'    Dim frmThis As Form
'    For Each frmThis In Forms
'        Unload frmThis
'    Next
End Sub

Public Property Get DeptID() As Long
    DeptID = glngExeDeptID
End Property

Public Property Let DeptID(ByVal vNewValue As Long)
    glngExeDeptID = vNewValue
End Property
'数据是否采集
Public Property Get DataReceived() As Boolean
    DataReceived = blnDataReceived
End Property

Public Sub Release()
'    Dim objEng As New JRO.JetEngine
    Dim objFileSys As New Scripting.FileSystemObject, strTmpFile As String
    
    
'    With gcnOracle
'        If .State = adStateOpen Then .Close
'    End With
'    With gcnAccess
'        If .State = adStateOpen Then .Close
'    End With
    Set gcnOracle = Nothing ': Set gcnAccess = Nothing

    '压缩Access数据库
'    On Error Resume Next
'    strTmpFile = App.Path & "\" & objFileSys.GetTempName
'    objEng.CompactDatabase "Data Source=" & App.Path & "\ZlLISLog.mdb", _
'        "Data Source=" & strTmpFile
'    objFileSys.CopyFile strTmpFile, App.Path & "\ZlLISLog.mdb"
'    objFileSys.DeleteFile strTmpFile, True
    
    Unload frmLISSrv
End Sub

'初始酶标仪参数
Public Sub InitMBPara()
    Dim strItem As String
    If gblnFromDB Then
        glngMBDeviceID = gobjDatabase.GetPara("酶标仪器", glngSys, 1208, -1)
        strItem = gobjDatabase.GetPara("酶标仪项目", glngSys, 1208, "")
        glngMBNo = gobjDatabase.GetPara("酶标仪标本号", glngSys, 1208, 1)
        gstrMBPosition = gobjDatabase.GetPara("酶标仪起始位置", glngSys, 1208, "A01")
    Else
        glngMBDeviceID = Val(GetSetting("ZLSOFT", "公共模块\zl9LISWork", "酶标仪器", -1))
        strItem = GetSetting("ZLSOFT", "公共模块\zl9LISWork", "酶标仪项目", "")
        glngMBNo = Val(GetSetting("ZLSOFT", "公共模块\zl9LISWork", "酶标仪标本号", ""))
        gstrMBPosition = GetSetting("ZLSOFT", "公共模块\zl9LISWork", "酶标仪起始位置", "A01")
    End If
    
    If Len(strItem) = 0 Then
        gstrMBChannel = ""
    Else
        gstrMBChannel = Split(strItem, "|")(1)
    End If
    

End Sub


