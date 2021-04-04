VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWardMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function CEC_Initialize Lib "CecDeviceToHis.dll" (ByVal ServerIP As String, ByVal ServerPort As Long, ByVal DispWindowhwnd As Long, ByVal recallFunAddr As Long, ByVal lngobj As Long) As Boolean
Private Declare Function CEC_Uninitialize Lib "CecDeviceToHis.dll" () As Boolean

Private Declare Function CEC_DevNo2His Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngType As Long, ByVal strInPatient As String) As Boolean
Private Declare Function CEC_His2DevNo Lib "CecDeviceToHis.dll" (ByVal lngType As Long, ByVal strInPatient As String) As Long
'lngType:1监护仪床号, 2HIS床号, 3病历编号

Private Declare Function CEC_ShowWindows Lib "CecDeviceToHis.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
Private Declare Function CEC_SelectBedNo Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long) As Boolean
Private Declare Function CEC_GetListBedNo Lib "CecDeviceToHis.dll" (ByVal strListBedNo As String) As Boolean

Private mblnUsable As Boolean

Public Function Init(ByVal cnOracle As ADODB.Connection, ByRef lngWindowHWND As Long, Optional ByVal lngWindowMode As Long = 3) As Boolean
'功能：初始化监护仪窗口
'参数：WindowHWND-显示监护仪的窗口句柄,WindowMode-显式模式
'        0-隐藏 1-普通  2-最大化，没有标题栏  3-最大化,有标题栏 4-最小化 5-隐藏床位下拉列表,增加退出按钮
    Dim ServerIP As String, ServerPort As String
    
    Call GetServerInfo(ServerIP, ServerPort)
    If ServerPort = "" Or ServerIP = "" Then GoTo errH
    
    On Error GoTo errH
    Call CEC_Initialize(ServerIP, ServerPort, lngWindowHWND, AddressOf RequestData, 0)
    '传入回调函数指针，用于接收数据请求
    
    Set gcnOracle = cnOracle
    
    Init = True
    mblnUsable = True
    Exit Function
errH:
    Init = False
    mblnUsable = False
End Function

Public Sub ShowInfor(ByVal strInPatient As String)
'功能：显示指定病人的监护信息
'参数：strInPatient-住院号
    Dim lngDevice As Long
        
    lngDevice = CEC_His2DevNo(3, strInPatient)
    If lngDevice = 0 Then
        Call ShowWindow(False)
    Else
        Call ShowWindow(True)
        Call CEC_SelectBedNo(lngDevice)
    End If
End Sub

Public Sub ShowWindow(ByVal blnShow As Boolean)
'功能：显示或隐藏监护窗口
'参数：blnShow:true显示，false隐藏
    Dim lngWindowMode As Long   '0-隐藏 1-普通  2-最大化，没有标题栏  3-最大化,有标题栏 4-最小化 5-隐藏床位下拉列表,增加退出按钮
    If blnShow Then
        lngWindowMode = 3
    Else
        lngWindowMode = 0
    End If
    Call CEC_ShowWindows(0, lngWindowMode)
End Sub

Public Sub GetListBedNo(ByRef strListBedNo As String)
'功能：获取现已连接到服务上的所有监护仪床号
    Dim strList As String * 4000, arrtmp As Variant, i As Long, lngDevice As Long, strPatiNO As String * 20
       
    Call CEC_GetListBedNo(strList)  '返回的是设备号列表 '返回格式：2|1,2，最前面的2是连接设备个数，'|'后再的为号列表
    If InStr(strList, "|") > 1 Then
        arrtmp = Split(strList, "|")
        For i = 1 To Val(arrtmp(0))
            lngDevice = Val(arrtmp(i))
            strPatiNO = ""
            Call CEC_DevNo2His(lngDevice, 3, strPatiNO)
            If Trim(strPatiNO) <> "" Then
                strListBedNo = IIf(strListBedNo = "", "", strListBedNo & ",") & Val(strPatiNO)
            End If
        Next
    End If
End Sub


Private Sub Class_Terminate()
    If mblnUsable Then Call CEC_Uninitialize
End Sub

Public Function ShowPatientInfor(ByVal strInPatient As String) As String
    ShowPatientInfor = GetPatientInfor(strInPatient)
End Function

Private Sub GetServerInfo(ByRef ServerIP As String, ByRef ServerPort As String)
'功能：获取配置文件中的服务器IP和端口
    Dim objFile As Object
    Dim objText As Object
    Dim strFile As String
    Dim strTmp As String
    
    On Error Resume Next
    Set objFile = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\zlWardMonitor.ini"
    Set objText = objFile.OpenTextFile(strFile, 1) '1-ForReading
    If Err.Number <> 0 Then
        Set objFile = Nothing
        Exit Sub
    End If
    
    strTmp = objText.ReadLine
    If InStr(strTmp, ":") > 0 Then ServerIP = Split(strTmp, ":")(1)
    strTmp = objText.ReadLine
    If InStr(strTmp, ":") > 0 Then ServerPort = Split(strTmp, ":")(1)
    
    If ServerIP = "" Or ServerPort = "" Then
        Call WriteLog("读取服务器配置文件zlWardMonitor.ini失败.")
    End If
    objText.Close
    Set objFile = Nothing
End Sub


Public Sub WriteLog(ByVal strInfo As String)
    '将调试信息写入文件中
    Dim objFile As Object
    Dim objText As Object
    Dim strFile As String
    
    On Error Resume Next
    Set objFile = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\zlWardMonitor.Log"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, 8) '8-ForAppending
    objText.WriteLine Now()
    objText.WriteLine strInfo
    objText.Close
    Set objFile = Nothing
End Sub
