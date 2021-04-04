Attribute VB_Name = "mdlQueueShow"
Option Explicit

'��Ϣ��Ƕ���
Public Const G_STR_MSG_QUEUE_001 As String = "ZLHIS_QUEUE_001" '�����Ϣ
Public Const G_STR_MSG_QUEUE_002 As String = "ZLHIS_QUEUE_002" '�����Ϣ
Public Const G_STR_MSG_QUEUE_003 As String = "ZLHIS_QUEUE_003" '״̬ͬ��
Public Const G_STR_MSG_QUEUE_004 As String = "ZLHIS_QUEUE_004" '��������

Public Enum TBusinessType
'ҵ�����Ͷ���
    btClinical = 0  '�ٴ��Ŷ�ҵ��
    btPacs = 1      'Pacs�Ŷ�ҵ��
    btPeis = 2       '����Ŷ�ҵ��
    'bt...          '���������ҵ�����ں��������չ
End Enum

Public Enum TShowStyle
'�Ŷӽкŵ���ʾ��ʽ
    ssSingleMan = 0     '��������ʽ
    ssSingleQueue = 1   '�����У���ȷ�����һ�ִ�м䣩������һ�����һ���ִ�м���ʾ
    ssMultiQueue = 2    '����У������ȷ�����һ�ִ�м�����򰴿����ŶӵĶ��У�,
    ssOld = 3           '�ϰ���ʾ
End Enum

Public Type TRect
'��ʾλ������
    lngLeft As Long         '������
    lngTop As Long          '��������
    lngWidth As Long        '��
    lngHeight As Long       '��
    lngMonitorIndex As Long '��ʾ������
End Type

Public Type TLcdCommonParameter
'LCDͨ�ò����ṹ
    ssShowStyle As TShowStyle           '������ʾ��ʽ

    lngCurDeptID As Long                '��ǰ����ID
    strCurDiagnoseRoom As String        '��ǰ��������
    
    strQueryQueueNames As String        '�������ƣ��������ʹ�á�,�����ŷָ�
    blnShowAdvertise As Boolean         '�Ƿ���ʾ���
    
    strFilter As String                 '��ʾ���ݹ�������
    lngCallingRows As Long
    lngQueueRows As Long
    
    blnConvertQueueName As Boolean      'ת�����ϰ�洢�����µĶ�������
    
    blnScrollDisplay As Boolean         '������ʾ
    blnFontAutoSizeToList As Boolean    '�����Զ���Ӧ�б�
    recPos  As TRect                    '������ʾ����
End Type

'ע��·��
Public Const G_STR_REGPATH = "����ģ��\zl9QueueShow"

Public gcnOracle As New ADODB.Connection    '�������ݿ�����

Public gobjStyleWindow() As Object
Private mobjIcon As clsTaskIcon

Public gobjComLib As Object            'zl9ComLib.clsComLib
Public gobjQueueShow As Object         'zl9LCDShow.clsLCDShow
Public gstrUserName As String
Public glngBusinessType As Long                     'LCD��ʾ����ҵ������
Public gstrSysName As String
Public gstrSystems As String
Public gstrStation As String
Public glngSys As String
Public gstrCompareVersion As String     '��ǰ�汾

Public gobjFile As New FileSystemObject

Public Sub Main()
    Dim objLogin As Object  'zlLogin.clsLogin
    Dim strCommand As String, strUserName As String, strPassword As String, strServer As String
    Dim blnAutoLogin As Boolean
    
    Set objLogin = DynamicCreate("zlLogin.clsLogin", "zlLogin.dll")
    If objLogin Is Nothing Then Exit Sub
    
    Set gobjComLib = DynamicGet("zl9ComLib.clsComLib", "zl9ComLib.dll")
    If gobjComLib Is Nothing Then Exit Sub
    
    If App.PrevInstance Then
        MsgBox "������ʾ�����Ѿ������������ٴ����С�", vbInformation, "����"
        Exit Sub
    End If
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls
    
    '�򿪵�½���棬���轫�û���������ע����У�����Ҫ����ע��·��
    
    blnAutoLogin = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "�Զ���¼", 0)) = 1
    
    If blnAutoLogin Then
        '��ע����л�ȡ��¼��Ϣ�����ܣ��Ա��Զ���½
        strUserName = getDecryptionPassW(GetSetting("ZLSOFT", G_STR_REGPATH, "�û���", ""))
        strPassword = getDecryptionPassW(GetSetting("ZLSOFT", G_STR_REGPATH, "����", ""))
        strServer = getDecryptionPassW(GetSetting("ZLSOFT", G_STR_REGPATH, "������", ""))
        
        If strUserName = "" Or strPassword = "" Or strServer = "" Then
            Set gcnOracle = objLogin.Login
        Else
            strCommand = "USER=" & strUserName & " PASS=" & strPassword & " SERVER=" & strServer
            Set gcnOracle = objLogin.Login(0, strCommand)
        End If
    Else
        Set gcnOracle = objLogin.Login
    End If
    
    If gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.ConnectionString = "" Then Exit Sub
    
    '�����½��Ϣ�����ܱ���
    SaveSetting "ZLSOFT", G_STR_REGPATH, "�û���", getEncryptionPassW(objLogin.InputUser)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "����", getEncryptionPassW(objLogin.InputPwd)
    SaveSetting "ZLSOFT", G_STR_REGPATH, "������", getEncryptionPassW(objLogin.ServerName)
    
    gstrSysName = "��ʾ"
    gstrUserName = objLogin.DBUser

    '��ʼ��zlcomlib����
    gobjComLib.InitCommon gcnOracle
    
    gstrSystems = " (ϵͳ =100 Or ϵͳ Is NULL)"
    glngSys = 100
    
    '������ָ������ʾҳ��
    Call ShowWindow(blnAutoLogin)
End Sub

Private Sub ShowWindow(ByVal blnAutoLogin As Boolean)
'�򿪶���ҳ��
'���ݲ�����ʾ��Ӧ�Ĵ���,��ʾ��������
'1.����������Զ���¼����ֱ���Դ�����ʽҳ�棬����������
'2.���û�������Զ���¼������ʾ���ô���
'���������ַ�ʽ��ʾ�󣬶���Ҫ��ʾ����ͼ��

On Error GoTo ErrorHand
    
    gstrCompareVersion = getCompareVersion
    
    Call InitOldLCDShow
    
    glngBusinessType = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "����ҵ��", 1))
    
    If blnAutoLogin Then
        '���ݲ�����ʾ��ʽ����
        Call OpenStyleWindow
    Else
        '��ʾ���ô���
        Call OpenMainCfg
    End If
    
    '������ͼ��
    Call OpenTrayIcon
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Function getCompareVersion() As String
'��ȡ��ǰϵͳ�汾
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
    getCompareVersion = ""
    
    strSql = "Select nvl(���汾,1) ���汾,nvl(�ΰ汾,0) �ΰ汾,nvl(���汾,0) ���汾,���� " & _
             "From ZlComponent Where Upper(Rtrim(����))=upper('zl9PacsWork') And ϵͳ=100"
             
    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ�汾��Ϣ")
    
    If rsRecord.RecordCount > 0 Then
        '��װ�汾��Ϊ��λ���汾����λ�ΰ汾����λ���汾
        getCompareVersion = String(3 - Len(rsRecord!���汾), "0") & rsRecord!���汾 & "." & _
                            String(3 - Len(rsRecord!�ΰ汾), "0") & rsRecord!�ΰ汾 & "." & _
                            String(3 - Len(rsRecord!���汾), "0") & rsRecord!���汾
    End If
End Function

Private Sub OpenOldLcd(ByVal lngShowNum As Long)
'���ϰ汾��LcdShow������ʾ
    Dim i As Integer
    Dim strQueueNames As String
    Dim str��������() As String     '���������谴�ϰ汾�ĸ�ʽ���룺��PACS:/*64:CT1,64:CT2....*/
    Dim blnConvertQueueName As Boolean  '�Ƿ�ת�����ϰ汾��ʽ�Ķ�������

    blnConvertQueueName = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & lngShowNum, "ת����������", 0)) = 1

    '����ҵ�����ͻ�ȡ��Ӧ��ʽ�Ķ�������
    strQueueNames = ConvertFormat(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & lngShowNum, "��ʾ����"), blnConvertQueueName)

    If strQueueNames = "" Then Exit Sub

    str�������� = Split(strQueueNames, ",")

    Call InitOldLCDShow

    Call gobjQueueShow.zlShow(gcnOracle, str��������, "", "", "", 0, False)
End Sub

Private Function ConvertFormat(ByVal strQueueName As String, ByVal blnConvertQueueName As Boolean) As String
'����ʽת�����Ѷ�������ת�������ݿ��д洢�ĸ�ʽ
    Dim i As Integer
    Dim str��������() As String, strQueueNames As String
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    Dim lngPreDeptID As Long
    Dim lngCurDeptID As Long
    Dim blnQueueStyle As Boolean
    
    If strQueueName = "" Then Exit Function
    
    strQueueNames = ""
    lngPreDeptID = 0
    lngCurDeptID = 0
    
    str�������� = Split(strQueueName, ",")
    
    If blnConvertQueueName Then    'ת�����ϰ汾��ʽ�Ķ�������
        For i = 0 To UBound(str��������)
            lngCurDeptID = Split(Split(str��������(i), "|")(1), "_")(0)
            
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    If lngPreDeptID <> lngCurDeptID Then
                        strQueueNames = strQueueNames & "," & lngCurDeptID
                    End If
                    
                Case TBusinessType.btPacs
                    If InStr(str��������(i), "���Ҷ���") Then
                        strQueueNames = strQueueNames & "," & Split(Split(str��������(i), "_")(1), ":")(0) & "-" & Split(str��������(i), "|")(0)
                    Else
                        strQueueNames = strQueueNames & "," & lngCurDeptID & ":" & Split(str��������(i), ":")(1)
                    End If
                    
                Case TBusinessType.btPeis
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        strQueueNames = strQueueNames & "," & Nvl(rsRecord!վ������) & ":" & Split(Split(str��������(i), "|")(1), ":")(1)
                    End If
                    
                'Case "" '''''
                '.
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next
    Else
        For i = 0 To UBound(str��������)
            lngCurDeptID = Split(Split(str��������(i), "|")(1), "_")(0)
            
            Select Case glngBusinessType
                Case TBusinessType.btClinical
                    If lngPreDeptID <> lngCurDeptID Then
                        strQueueNames = strQueueNames & "," & lngCurDeptID
                    End If
                    
                Case TBusinessType.btPacs
                    strQueueNames = strQueueNames & "," & Split(str��������(i), "|")(0) & "-" & Split(str��������(i), ":")(1)
                    
                Case TBusinessType.btPeis
                    strSql = "select վ������ from ���վ��ֲ� where ִ�п���id=[1]"
                    Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡվ������", lngCurDeptID)
                    
                    If rsRecord.RecordCount > 0 Then
                        strQueueNames = strQueueNames & "," & Nvl(rsRecord!վ������) & ":" & Split(str��������(i), ":")(1)
                    End If
                    
                'Case "" '''''
                '.
            End Select
            
            lngPreDeptID = lngCurDeptID
        Next
    End If
    
    ConvertFormat = strQueueNames
End Function

Private Sub OpenMainCfg()
'�������ô���
    Call frmMain.zlShowMe
End Sub

Public Sub OpenStyleWindow()
'�������ô�����ʽ���ڶ��󣬲��򿪶�Ӧ����ʽ������ʾ
'blnOpenOldLcd,�Ƿ����ϰ��LCDģʽ��ʾ
    Dim i As Integer
    Dim lngShowNum As Long
    Dim strShowStyle As String

    lngShowNum = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��������", 1))
    
    ReDim gobjStyleWindow(lngShowNum) As Object
    
    For i = 1 To lngShowNum
        strShowStyle = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & i, "��ʾ��ʽ", "1-��������ʽ")
        
        Select Case Split(strShowStyle, "-")(0)
            Case TShowStyle.ssSingleMan
                Set gobjStyleWindow(i) = New frmStyle_SingleMan

            Case TShowStyle.ssSingleQueue
                Set gobjStyleWindow(i) = New frmStyle_SingleQueue
                
            Case TShowStyle.ssMultiQueue
                Set gobjStyleWindow(i) = New frmStyle_MultiQueue

            Case TShowStyle.ssOld
                Set gobjStyleWindow(i) = Nothing
            
            'Case TShowStyle.ssOther...
            '    Set gobjStyleWindow(i) = New frmStyle_Other...
            '
            '...
        End Select
        
        If Split(strShowStyle, "-")(0) = TShowStyle.ssOld Then
            Call OpenOldLcd(i)
        Else
            Call gobjStyleWindow(i).ISty_Show(i)
        End If
    Next
End Sub

Public Sub CloseStyleWindow()
'�ر�������ʽ����
    Dim i As Integer
    '������õ����ϰ��LCDSHOW��ر��ϰ��LCDSHOW
    If Not gobjQueueShow Is Nothing Then
        gobjQueueShow.zlclose
        Set gobjQueueShow = Nothing
    End If
    
    If SafeArrayGetDim(gobjStyleWindow) <= 0 Then Exit Sub
    
    For i = 1 To UBound(gobjStyleWindow)
        If Not gobjStyleWindow(i) Is Nothing Then Unload gobjStyleWindow(i)
    Next
End Sub

Private Sub OpenTrayIcon()
'������ͼ��
    frmTrayIcon.Show
    frmTrayIcon.Hide
End Sub

Public Sub InitOldLCDShow()
'��ʼ���ϰ汾LCD��ʾ����
    If gobjFile.FileExists("C:\APPSOFT\Apply\zl9LCDShow.dll") Then
        Set gobjQueueShow = DynamicCreate("zl9LCDShow.clsLCDShow", "zl9LCDShow.dll")
    End If
End Sub

