VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mobjJson As clsJson
Private mfrmMain As frmMain

Public Function UserLogin(ByVal strIP As String, _
    ByVal strPort As String, ByVal strSID As String, _
    ByVal strUser As String, ByVal strPwd As String) As Boolean
'���ܣ��û���ʼ����¼�����û���¼�л�ʱ��Ҫ���øú���
    Dim objRegister As Object
    Dim objComlib As Object
    Dim strError As String
    
    Call ResetConnection
    
    '����zlRegister���󣺾ֲ�����ÿ�ζ�̬������Ϊ��ʵ�ֲ����ļ�ռ��
    On Error Resume Next
    Set objRegister = CreateObject("zlRegister.clsRegister")
    If objRegister Is Nothing Then
        Err.Clear: Exit Function
    End If
    Err.Clear: On Error GoTo 0
    
    '�������ӣ�ʹ��OraOLEDB
    Set gcnOracle = objRegister.GetConnection(strIP & ":" & strPort & "/" & strSID, strUser, strPwd, True, 1, strError)
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State = adStateClosed Then
        Set gcnOracle = Nothing
        Exit Function
    End If
    
    Set objRegister = Nothing '�ļ�ռ���ͷŲ��ˣ�
    
    '��ʾ������Ϊ����ȷ��λ��
    If mfrmMain Is Nothing Then
        Set mfrmMain = New frmMain
        mfrmMain.Show: mfrmMain.Hide
    End If
    
    Call InitComLib '�漰�����л�����Ϊ��������õ�����������Ҫ��ʼ����������
    Call InitPlugIn '�漰�����л�
    Call InitReport
    
    UserLogin = True
End Function

Public Function ClinicBillPrintSetup(ByVal strReportNo As String) As Boolean
'���ܣ����Ƶ��ݴ�ӡ����
    If gobjReport Is Nothing Then Exit Function
    
    gstrHwndOLD = "": EnumChildWindows GetDesktopWindow, AddressOf EnumChildProcOld, ByVal 0
    If Not mfrmMain Is Nothing Then
        mfrmMain.TimerShow.Enabled = True
    End If
    
    ClinicBillPrintSetup = gobjReport.ReportPrintSet(gcnOracle, 100, strReportNo, mfrmMain)
End Function

Public Function ClinicBillPreview(ByVal strReportNo As String, ByVal strBillNo As String) As Boolean
'���ܣ�����Ԥ��
    If gobjReport Is Nothing Then Exit Function
    
    gstrHwndOLD = "": EnumChildWindows GetDesktopWindow, AddressOf EnumChildProcOld, ByVal 0
    If Not mfrmMain Is Nothing Then
        mfrmMain.TimerShow.Enabled = True
    End If
    
    ClinicBillPreview = gobjReport.ReportOpen(gcnOracle, 100, strReportNo, mfrmMain, "NO=" & strBillNo, "����=1", "ҽ��ID=0", "PrintEmpty=0", 1)
End Function

Public Function ClinicBillPrint(ByVal strJson As String) As Boolean
'���ܣ�������ӡ
'������Json��ʽ={"bills":[{"report_no":"xxx","bill_no":"xxx"},...]}
    Dim lngLen As Long, i As Long
    Dim strReportNo As String, strBillNo As String
    
    If strJson = "" Then Exit Function
    If gobjReport Is Nothing Then Exit Function

    If mobjJson Is Nothing Then
        Set mobjJson = New clsJson
    End If
    Call mobjJson.OpenJson(strJson)
    
    lngLen = mobjJson.GetLength("bills")
    For i = 0 To lngLen - 1
        strReportNo = mobjJson.GetValue("bills[" & i & "].report_no")
        strBillNo = mobjJson.GetValue("bills[" & i & "].bill_no")
        
        Call gobjReport.ReportOpen(gcnOracle, 100, strReportNo, mfrmMain, "NO=" & strBillNo, "����=1", "ҽ��ID=0", "PrintEmpty=0", "SilentMode=1", 2)
    Next
    
    ClinicBillPrint = True
End Function

Public Function GetPlugInFuncs() As String
'���ܣ���ȡ����ҽ��վ�Ĳ����������
    
    If gobjPlugIn Is Nothing Then Exit Function
    
    '�����Ǽ���ZLHIS������ҽ��վ������ʹ��1252
    GetPlugInFuncs = gobjPlugIn.GetFuncNames(100, 1252, 0)
End Function

Public Function ExecPlugInFuncs(ByVal strFuncName As String, ByVal lngPatiID As Long, _
    ByVal lngVisitID As Long, ByVal strKey As String, Optional ByVal strReserve As String) As Boolean
'���ܣ�ִ������ҽ��վ�Ĳ������
'������
'  strKey=PatiAdviceRecord��ID��Ϊ36λ�ַ�����GUID
'  strReserve=��ǰѡ����Ŀ��"�������:��ĿID(������ĿID��ҩƷID)"
    
    If gobjPlugIn Is Nothing Then Exit Function
    
    gstrHwndOLD = "": EnumChildWindows GetDesktopWindow, AddressOf EnumChildProcOld, ByVal 0
    If Not mfrmMain Is Nothing Then
        mfrmMain.TimerShow.Enabled = True
    End If
    
    '�����Ǽ���ZLHIS������ҽ��վ������ʹ��1252
    Call gobjPlugIn.ExecuteFunc(100, 1252, strFuncName, lngPatiID, lngVisitID, strKey, strReserve, 0)
    
    ExecPlugInFuncs = True
End Function

Private Function InitComLib() As Boolean
'���ܣ���ʼ����������
    If gobjComLib Is Nothing Then
        On Error Resume Next
        Set gobjComLib = CreateObject("zl9ComLib.clsComlib")
        If gobjComLib Is Nothing Then
            Err.Clear: Exit Function
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    '�����л�����
    Call gobjComLib.InitCommon(gcnOracle)
    
    InitComLib = True
End Function

Private Function InitPlugIn() As Boolean
'���ܣ���ʼ�����
'������blnForce=�Ƿ�ǿ�Ƴ�ʼ���������û����µ�¼��
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If gobjPlugIn Is Nothing Then
            Err.Clear: Exit Function
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    '�����Ǽ���ZLHIS������ҽ��վ������ʹ��1252��
    InitPlugIn = gobjPlugIn.Initialize(gcnOracle, 100, 1252, 0) '�е���
End Function

Private Function InitReport() As Boolean
'���ܣ���ʼ����������
    If gobjReport Is Nothing Then
        On Error Resume Next
        Set gobjReport = CreateObject("zl9Report.clsReport")
        If gobjReport Is Nothing Then
            Err.Clear: Exit Function
        End If
        Err.Clear: On Error GoTo 0
    End If
    
    InitReport = True
End Function

Private Sub ResetConnection()
'���ܣ����Ӹ�λ
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            gcnOracle.Close
        End If
    End If
    Set gcnOracle = Nothing
End Sub

Private Sub Class_Terminate()
    If Not mobjJson Is Nothing Then
        Set mobjJson = Nothing
    End If
    If Not mfrmMain Is Nothing Then
        Unload mfrmMain
        Set mfrmMain = Nothing
    End If
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Terminate(100, 1252, 0)
        Set gobjPlugIn = Nothing
    End If
    
    If Not gobjReport Is Nothing Then
        Set gobjReport = Nothing
    End If
    
    If Not gobjComLib Is Nothing Then
        Set gobjComLib = Nothing
    End If
    
    Call ResetConnection
End Sub