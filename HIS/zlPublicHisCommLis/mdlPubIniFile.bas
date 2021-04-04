Attribute VB_Name = "mdlPubIniFile"
Option Explicit

Public gcollLog As New Collection   '��־����

'��дini �ļ���API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'����Ŀ¼
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp" (ByVal PathName As String) As Long

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = Replace(GetStr, Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function


Public Sub SaveDBLog(intSource As Integer, intPriority As Integer, lngSampleID As Long, strProperties As String, strExplain As String, Optional intNO As Integer, Optional strProgram As String, Optional lngMachineID As Long)
          '����     ������־
          '����     intSource       ��Դ
          '         intPriority    ���ȼ�
          '         lngSampleID    �걾ID
          '         strProperties  ����
          '         strExplain     ˵��
          '         intNo          ģ���
          '         strProgram     ģ������
          Dim strSQL As String
          Dim strLocalIP As String

1     On Error GoTo SaveDBLog_Error

'2         strLocalIP = frmCommSetup.Winsock.LocalIP
          strLocalIP = gobjHisSystem.IP
3         strSQL = "Zl_���������־_insert(" & intSource & "," & intPriority & ",'" & gUserInfo.Name & "'," & IIf(lngSampleID = 0, "null", lngSampleID) & ",'" & strProperties & "','" & strExplain & "'," & intNO & ",'" & strProgram & "'," & IIf(lngMachineID = 0, "null", lngMachineID) & ",'" & strLocalIP & "')"
4         Call ComExecuteProc(Sel_Lis_DB, strSQL, "��־����")

5         Exit Sub
SaveDBLog_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "ִ��(SaveDBLog)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/7/12
'��    ��:������־
'��    ��:
'           strObject       ����Ĺ�����
'           strModule       �����ģ��
'           strTxt          ������Ϣ
'           [blnMsg         �Ƿ񵯳���ʾ��]
'           [intType        0=������־��1=ͨѶ��־]
'��    ��:
'��    ��:  ������ʾ
'---------------------------------------------------------------------------------------
Public Function WriteErrLog(ByVal strObject As String, ByVal strModule As String, ByVal strErrTxt As String, _
                            Optional ByVal blnMsg As Boolean, Optional ByVal intType As Integer) As String
          Dim objFSO As New FileSystemObject
          Dim objFolder As Folder
          Dim objFiles As Files
          Dim objFile As File
          Dim strFileName As String
          Dim dFileCreatDate As Date
          Dim objTxt As TextStream
          Dim strFolder As String         '������־Ŀ¼
          Dim strFilePath As String       '������־·��
          Dim dTimeNow As Date            '��ǰϵͳʱ��
          Dim strDrive As String          '��������
          Dim DrDrive As Drive            '����������
          Dim strOrcMsg As String

1         On Error GoTo Errhand

2         strFolder = Mid(App.Path, 1, InStrRev(App.Path, "\")) & "PUBLIC"
3         If intType = 0 Then
4             strFolder = strFolder & "\Log\������־\"
5             strFilePath = strFolder & "������־.log"
6         Else
7             strFolder = strFolder & "\Log\ͨѶ��־\"
8             strFilePath = strFolder & "ͨѶ��־.log"
9         End If
10        dTimeNow = Now    ' ʹ��ComCurrDateʱ�������ComCurrDate�ڲ����ִ�����ComCurrDate�ڲ����ô�����־��writeErrLog������������ѭ��

          '���Ŀ¼�Ƿ���ڣ��������򴴽�Ŀ¼
11        If Not objFSO.FolderExists(strFolder) Then
12            Call MakeSureDirectoryPathExists(strFolder)
13        End If

          '����ļ��Ƿ���ڣ��������򴴽�
14        If objFSO.FileExists(strFilePath) = False Then
              'ɾ����������30��ı����ļ�
15            Set objFolder = objFSO.GetFolder(strFolder)
16            Set objFiles = objFolder.Files
17            For Each objFile In objFiles
18                strFileName = objFile.Name
19                If IsDate(Format(Mid(strFileName, 4, 4) & "-" & Mid(strFileName, 8, 2) & "-" & Mid(strFileName, 10, 2), "yyyy-mm-dd")) Then
20                    dFileCreatDate = CDate(Format(Mid(strFileName, 4, 4) & "-" & Mid(strFileName, 8, 2) & "-" & Mid(strFileName, 10, 2), "yyyy-mm-dd"))
21                    If DateDiff("d", dFileCreatDate, dTimeNow) > 30 Then
22                        objFSO.DeleteFile (objFile.Path)
23                    End If
24                End If
25            Next

              '������־�ļ�֮ǰ�ȼ���Ƿ����㹻�Ĵ��̿ռ�
26            strDrive = objFSO.GetDriveName(objFSO.GetAbsolutePathName(strFolder))    '��ȡ��־�ļ����ڵĴ�������
27            Set DrDrive = objFSO.GetDrive(strDrive)    '��ȡ���̶���
28            If DrDrive.IsReady Then    '���ô����Ƿ����
29                If DrDrive.FreeSpace < 52428800 * 2 Then    '�������ʣ��ռ�С��100M�����ֹ������־�ļ�
30                    MsgBox Mid(strDrive, 1, 1) & "�̿ռ䲻�㣬�޷�������־", vbInformation, gSysInfo.AppName
31                    Exit Function
32                End If
33            End If
34            objFSO.CreateTextFile (strFilePath)    '������־�ļ�
35        Else
              '����ļ��Ѿ����������ļ���С���ļ�̫��Ӱ���ȡЧ�ʣ������޶�Ϊ50M������50M�����ɱ����ļ�
36            Set objFile = objFSO.GetFile(strFilePath)
37            If objFile.Size > 52428800 Then    '�ж��ļ��Ƿ����50M
38                Name strFilePath As strFolder & "\LIS" & Format(dTimeNow, "yyyymmddhhmm") & ".bak"   '�޸��ļ���
                  '���ļ���������֮����Ҫ���¼�鲢�����ļ�
39                Call WriteErrLog(strObject, strModule, strErrTxt)
40                Exit Function
41            End If
42        End If

          'д�������־
43        On Error Resume Next
44        Set objTxt = objFSO.OpenTextFile(strFilePath, ForAppending, False)
45        objTxt.WriteLine "====================" & strObject & ":" & strModule & " " & dTimeNow & "=================="    '������ĸ������Ǹ�ģ����Ĵ�
46        objTxt.WriteLine strErrTxt
47        objTxt.Close
48        Set objTxt = Nothing

49        On Error GoTo Errhand

50        WriteErrLog = strErrTxt

          '�洢�����е�������ʾ
51        If InStr(UCase(strErrTxt), "[ZLSOFT]") > 0 Then
52            strOrcMsg = Mid(UCase(strErrTxt), InStr(UCase(strErrTxt), "[ZLSOFT]") + 8, InStrRev(UCase(strErrTxt), "[ZLSOFT]") - InStr(UCase(strErrTxt), "[ZLSOFT]") - 8)
53            WriteErrLog = strOrcMsg
54        End If

55        If blnMsg And intType = 0 Then
56            If strOrcMsg <> "" Then
                  '�洢�����е�������ʾ
57                MsgBox strOrcMsg, vbInformation, gSysInfo.AppName
58            Else
59                MsgBox strObject & ":" & strModule & ":" & strErrTxt, vbInformation, gSysInfo.AppName
60            End If
61        End If


62        Exit Function
Errhand:
63        MsgBox "д����־����,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, vbInformation, "��ʾ"
64        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2020-02-12
'��    ��:  ͨ����־���
'��    ��:
'           Oracle          ���Ӷ���
'           blnCallEnd      �Ƿ��ǽ���ʱ���ã�True=��
'           strProcName     ���÷�����
'           strTitle        ����
'           arrProcParas    ����
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Sub ExportLog(ByVal intSelDB As Integer, ByVal blnCallEnd As Boolean, ByVal strProcName As String, _
                     ByVal strTitle As String, ByVal strSQL As String, ParamArray arrProcParas() As Variant)
                     
          Const STR_CATEGORY As String = "������־"
          Dim strServerName As String
          Dim objLog As Object
          Dim arrPars() As Variant
          Dim Oracle As ADODB.Connection

1         On Error GoTo ExportLog_Error

2         If intSelDB = Sel_Lis_DB Then
3             Set Oracle = gcnLisOracle
4         ElseIf intSelDB = Sel_His_DB Then
5             Set Oracle = gcnHisOracle
6         End If

          '������־����
7         Call CreateLog(Oracle)

8         strServerName = GetServerName(Oracle)

9         On Error Resume Next
10        Set objLog = gcollLog(strServerName)
11        Err.Clear: On Error GoTo ExportLog_Error
12        If objLog Is Nothing Then
13            Exit Sub
14        End If

15        arrPars = arrProcParas
16        If strSQL = "" Then
17            Call objLog.LogCall(STR_CATEGORY, App.ProductName, "clsDatabase", strProcName, strTitle _
                  , IIf(blnCallEnd, Val("1-����"), Val("0-��ʼ")), arrPars)
18        Else
19            Call objLog.LogCall(STR_CATEGORY, App.ProductName, "clsDatabase", strProcName, strTitle _
                  , IIf(blnCallEnd, Val("1-����"), Val("0-��ʼ")), strSQL, arrPars)
20        End If


21        Exit Sub
ExportLog_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "ִ��(ExportLog)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2020-02-12
'��    ��:  ������־����
'��    ��:
'           Oracle      ���Ӷ���
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function CreateLog(Oracle As ADODB.Connection) As Boolean
          Dim objLog As Object
          Dim strServerName As String
          Dim i As Integer

1         On Error GoTo CreateLog_Error

2         strServerName = GetServerName(Oracle)

          '������־����
3         If gcollLog.Count = 0 Then
              '��һ�ν�����Ϊ�գ��Զ�����
4             On Error Resume Next
5             Set objLog = CreateObject("zlLog.clsLog")
6             Err.Clear: On Error GoTo CreateLog_Error
7             Call objLog.SetBusinessDB(Oracle)
8             gcollLog.Add objLog, strServerName
9         Else
              '��鵱ǰ���Ӷ����Ƿ񴴽��˶�Ӧ����־���������û�У��򴴽�
10            On Error Resume Next
11            Set objLog = gcollLog(strServerName)
12            Err.Clear: On Error GoTo CreateLog_Error
13            If objLog Is Nothing Then
14                On Error Resume Next
15                Set objLog = CreateObject("zlLog.clsLog")
16                Err.Clear: On Error GoTo CreateLog_Error
17                Call objLog.SetBusinessDB(Oracle)
18                gcollLog.Add objLog, strServerName
19            End If
20        End If


21        Exit Function
CreateLog_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "ִ��(CreateLog)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2020-02-12
'��    ��:  ͨ�����Ӷ����ȡ��������
'��    ��:
'           Oracle      ���Ӷ���
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function GetServerName(Oracle As ADODB.Connection) As String
          Dim strServerName As String
          Dim lngS As Long
          Dim lngE As Long

1         On Error GoTo GetServerName_Error

2         lngS = InStr(UCase(Oracle), "SOURCE=") + Len("SOURCE=")
3         lngE = InStr(lngS + 1, Oracle, ";")
4         strServerName = Mid(Oracle, lngS, lngE - lngS)
5         If InStr(strServerName, """(") > 0 Then
6             lngS = InStr(UCase(Oracle), "SERVICE_NAME=") + Len("SERVICE_NAME=")
7             lngE = InStr(lngS + 1, Oracle, ")")
8             strServerName = Mid(Oracle, lngS, lngE - lngS)
9         End If

10        GetServerName = strServerName


11        Exit Function
GetServerName_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlPubIniFile", "ִ��(GetServerName)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear

End Function
