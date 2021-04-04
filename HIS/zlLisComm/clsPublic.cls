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
    
'    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
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
'---  ��ʼ����������


    '�����־�ļ�
    If gblnFromDB Then
        gblnClearData = gobjDatabase.GetPara("��ս�����־", glngSys, 1208, 1)
    Else
        gblnClearData = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", 1)
    End If
    If gblnClearData Then
        On Error GoTo DBError
        If Dir(App.Path & "\zlLis*.LOG") <> "" Then
            Kill App.Path & "\zlLis*.LOG"
        End If
        
        On Error GoTo 0
    End If
    
    '��ʼø���ǲ���

    InitMBPara

    '-------------------------------------------------
    frmLISSrv.Show
    Exit Sub
DBError:
    Call WriteLog("clsPublic", "InitClsPublic", LOG_������־, Err.Description)
End Sub

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
'���ͱ걾��¼������
'    frmLISSrv.Show
    SendSample = frmLISSrv.SendSample(lngDeviceID, strSampleDate, strSampleNO, strAdviceIDs, blnUndo, iType)
End Function

Public Sub GetResultFromFile(ByVal strFile As String, ByVal lngDeviceID As Long, ByVal strSampleNO As String, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))
'���ļ���ȡ������
'   strFile������·�����ļ���
'   lngDeviceID�������豸ID
'   strSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��strSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��

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
'�����Ƿ�ɼ�
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

    'ѹ��Access���ݿ�
'    On Error Resume Next
'    strTmpFile = App.Path & "\" & objFileSys.GetTempName
'    objEng.CompactDatabase "Data Source=" & App.Path & "\ZlLISLog.mdb", _
'        "Data Source=" & strTmpFile
'    objFileSys.CopyFile strTmpFile, App.Path & "\ZlLISLog.mdb"
'    objFileSys.DeleteFile strTmpFile, True
    
    Unload frmLISSrv
End Sub

'��ʼø���ǲ���
Public Sub InitMBPara()
    Dim strItem As String
    If gblnFromDB Then
        glngMBDeviceID = gobjDatabase.GetPara("ø������", glngSys, 1208, -1)
        strItem = gobjDatabase.GetPara("ø������Ŀ", glngSys, 1208, "")
        glngMBNo = gobjDatabase.GetPara("ø���Ǳ걾��", glngSys, 1208, 1)
        gstrMBPosition = gobjDatabase.GetPara("ø������ʼλ��", glngSys, 1208, "A01")
    Else
        glngMBDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\zl9LISWork", "ø������", -1))
        strItem = GetSetting("ZLSOFT", "����ģ��\zl9LISWork", "ø������Ŀ", "")
        glngMBNo = Val(GetSetting("ZLSOFT", "����ģ��\zl9LISWork", "ø���Ǳ걾��", ""))
        gstrMBPosition = GetSetting("ZLSOFT", "����ģ��\zl9LISWork", "ø������ʼλ��", "A01")
    End If
    
    If Len(strItem) = 0 Then
        gstrMBChannel = ""
    Else
        gstrMBChannel = Split(strItem, "|")(1)
    End If
    

End Sub

