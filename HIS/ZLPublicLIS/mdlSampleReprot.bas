Attribute VB_Name = "mdlSampleReprot"
Option Explicit

Public gstrSysName As String                        'ϵͳ����
Public gstrProductName As String                    'OEM��Ʒ����
Public gstrUnitName As String                       '�û���λ����
Public gcnOracle As New ADODB.Connection                 '�������ݿ�����

Public UserInfo As TYPE_USER_INFO

'�û���Ϣ
Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String '��Ա����
    ���� As String
    DeptID As Long '����ID
    DeptNo As String '���ű��
    DeptName As String '��������
    DBUser As String '���ݿ��û�
End Type

Public glngSys As Long                              'ϵͳ��
Public glngModule As Long                           'ģ���
Public gobjLISInsideComm As Object
Public gobjComLib As Object
Public gobjFile As New FileSystemObject

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.DeptID = Nvl(rsTmp!����ID, 0)
            UserInfo.DeptNo = rsTmp!������ & ""
            UserInfo.DeptName = rsTmp!������ & ""
            UserInfo.DBUser = rsTmp!�û��� & ""
            GetUserInfo = True
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub InitObjLis()
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLISInsideComm Is Nothing Then
        On Error Resume Next
        Set gobjLISInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLISInsideComm Is Nothing Then
            If gobjLISInsideComm.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLISInsideComm = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Public Function writeErrLog(ByVal strObject As String, ByVal strModule As String, ByVal strErrTxt As String, Optional ByVal blnMsg As Boolean) As Boolean
    '������־
    'strObject      ����Ĺ�����
    'strModule      �����ģ��
    'strTxt         ��־����
    
    Dim objFSO As New FileSystemObject
    Dim objFolder As Folder
    Dim objFiles As Files
    Dim objFile As File
    Dim strFilename As String
    Dim dFileCreatDate As Date
    Dim objTxt As TextStream
    Dim strFolder As String         '������־Ŀ¼
    Dim strFilePath As String       '������־·��
    Dim dTimeNow As Date            '��ǰϵͳʱ��
    Dim strDrive As String          '��������
    Dim DrDrive As Drive            '����������
    
    On Error GoTo errhand
    
'    strFolder = App.Path & "\LisErrLog"
'    strFilePath = App.Path & "\LisErrLog\LisErrLog.log"
'    dTimeNow = Now 'ComCurrDate
'
'
'    '���Ŀ¼�Ƿ���ڣ��������򴴽�Ŀ¼
'    If Not objFSO.FolderExists(strFolder) Then
'        objFSO.CreateFolder (strFolder)
'    End If
'
'    '����ļ��Ƿ���ڣ��������򴴽�
'     If objFSO.FileExists(strFilePath) = False Then
'        'ɾ����������30��ı����ļ�
'        Set objFolder = objFSO.GetFolder(strFolder)
'        Set objFiles = objFolder.Files
'        For Each objFile In objFiles
'            strFilename = objFile.Name
'            If IsDate(Format(Mid(strFilename, 4, 4) & "-" & Mid(strFilename, 8, 2) & "-" & Mid(strFilename, 10, 2), "yyyy-mm-dd")) Then
'                dFileCreatDate = CDate(Format(Mid(strFilename, 4, 4) & "-" & Mid(strFilename, 8, 2) & "-" & Mid(strFilename, 10, 2), "yyyy-mm-dd"))
'                If DateDiff("d", dFileCreatDate, dTimeNow) > 30 Then
'                    objFSO.DeleteFile (objFile.Path)
'                End If
'            End If
'        Next
'
'        '������־�ļ�֮ǰ�ȼ���Ƿ����㹻�Ĵ��̿ռ�
'        strDrive = objFSO.GetDriveName(objFSO.GetAbsolutePathName(strFolder)) '��ȡ��־�ļ����ڵĴ�������
'        Set DrDrive = objFSO.GetDrive(strDrive) '��ȡ���̶���
'        If DrDrive.IsReady Then '���ô����Ƿ����
'            If DrDrive.FreeSpace < 52428800 * 2 Then    '�������ʣ��ռ�С��100M�����ֹ������־�ļ�
'                MsgBox Mid(strDrive, 1, 1) & "�̿ռ䲻�㣬�޷�������־", vbInformation, "��ʾ"
'                Exit Function
'            End If
'        End If
'        objFSO.CreateTextFile (strFilePath) '������־�ļ�
'     Else
'        '����ļ��Ѿ����������ļ���С���ļ�̫��Ӱ���ȡЧ�ʣ������޶�Ϊ50M������50M�����ɱ����ļ�
'        Set objFile = objFSO.GetFile(strFilePath)
'        If objFile.Size > 52428800 Then '�ж��ļ��Ƿ����50M
'            Name strFilePath As strFolder & "\LIS" & Format(dTimeNow, "yyyymmddhhmm") & ".bak"   '�޸��ļ���
'            '���ļ���������֮����Ҫ���¼�鲢�����ļ�
'            Call writeErrLog(strObject, strModule, strErrTxt)
'        End If
'    End If
'
'    'д�������־
'    Set objTxt = objFSO.OpenTextFile(strFilePath, ForAppending, False)
'    objTxt.WriteLine "====================" & strObject & ":" & strModule & " " & dTimeNow & "==================" '������ĸ������Ǹ�ģ����Ĵ�
'    objTxt.WriteLine strErrTxt
'    objTxt.Close
'    Set objTxt = Nothing
    
    If blnMsg Then MsgBox "��Ǹ,������ʹ�õĹ��ܳ����쳣,�뼰ʱ��ϵ����ṩ��", vbInformation, "��ʾ"
    
    '���ù���������¼��־
    Call zl9ComLib.LogWrite("������־", strModule, Mid(strErrTxt, 4, InStr(strErrTxt, ")") - 4), strErrTxt)
    
    writeErrLog = True
    Exit Function
errhand:
    writeErrLog = False
    MsgBox "д�������־����,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, vbInformation, "��ʾ"
    Err.Clear
End Function


Public Function ComSetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '���ò���
    ComSetPara = zlDatabase.SetPara(varPara, strValue, lngSys, lngModual, blnSetup)
End Function

Public Function ComGetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
    'ȡ����
    ComGetPara = zlDatabase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
End Function
