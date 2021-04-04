Attribute VB_Name = "mImage"
Option Explicit

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_HIDDEN = &H2


Public Const IMG_LAB_CHECKBOX_TAG = "CHECKBOX"
Public Const IMG_LAB_HINT_TAG = "HINT"
Public Const IMG_LAB_ORDER_TAG = "ORDER"
Public Const IMG_LAB_ERRORINFO_TAG = "ERRORINFO"
Public Const IMG_LAB_ERRORSTATE_TAG = "ERRORSTATE"

Public Const IMG_BACK_BORDER_COLOR = &HE0E0E0

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As String, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private mlastDicomInfo As TDicomBaseInfo
Private mlineFtpInfo As TFtpDeviceInf
Private mbackFtpInfo As TFtpDeviceInf


Public Function IsExistsBGServer() As String
'����̨��������ļ��Ƿ����
    Dim strServerFile As String
    Dim strBgExe As String
    
    IsExistsBGServer = ""
    
    strBgExe = "ZL9PACSIMGTRANS"
    
    strServerFile = FormatFilePath(SysRootPath & "\" & strBgExe & ".exe")
    If Trim(Dir(strServerFile, vbHidden)) <> "" Then
        IsExistsBGServer = strServerFile
        Exit Function
    End If
        
    strServerFile = FormatFilePath(SysRootPath & "\Apply\" & strBgExe & ".exe")
    If Trim(Dir(strServerFile, vbHidden)) <> "" Then
        IsExistsBGServer = strServerFile
        Exit Function
    End If
        
    strServerFile = FormatFilePath(SysRootPath & "\PUBLIC\" & strBgExe & ".exe")
    If Trim(Dir(strServerFile, vbHidden)) <> "" Then
        IsExistsBGServer = strServerFile
        Exit Function
    End If
End Function

Public Function GetBgImgInfo(dcmInfo As TDicomBaseInfo, _
    lineFtpInfo As TFtpDeviceInf, backFtpInfo As TFtpDeviceInf, _
    Optional ByVal blnIsUpload As Boolean = True) As clsBgImgInfo
    
    Dim objBgImgInfo As clsBgImgInfo
    
    
    Set objBgImgInfo = New clsBgImgInfo
    
    objBgImgInfo.Key = dcmInfo.strInstanceUID
    objBgImgInfo.Filename = dcmInfo.strInstanceUID
    objBgImgInfo.FilePath = GetStudyImgPath(dcmInfo)  'ͼ�����ڱ��ش洢·��
    objBgImgInfo.StudyUID = dcmInfo.strStudyUID
    
    objBgImgInfo.AdviceId = dcmInfo.lngAdviceId
    objBgImgInfo.PatientName = dcmInfo.strName
    objBgImgInfo.SeriesNoTag = dcmInfo.lngSeriesNo
    
    If blnIsUpload Then
        objBgImgInfo.ImgCommand = icUpLoad
        
        If dcmInfo.lngMediaTag = 0 Then objBgImgInfo.JpgConvert = True
    Else
        objBgImgInfo.ImgCommand = icReadly
    End If
    
    
    objBgImgInfo.IsBackGround = True '�Ӳ�����ȡ�Ƿ��̨����ͼ���ϴ�
    
    Select Case dcmInfo.lngMediaTag
        Case ImgTag
            objBgImgInfo.Format = ifDcm
        Case VIDEOTAG
            objBgImgInfo.Format = ifAvi
        Case AUDIOTAG
            objBgImgInfo.Format = ifWav
    End Select
    
    With lineFtpInfo
        objBgImgInfo.FtpIp = .strFtpIp
        objBgImgInfo.FtpUser = .strFTPUser
        objBgImgInfo.FtpPwd = .strFTPPwd
        objBgImgInfo.FtpVirtualPath = .strFtpVirtualURL
        objBgImgInfo.FtpFile = dcmInfo.strInstanceUID
    End With

    If Len(backFtpInfo.strDeviceId) > 0 Then
    With backFtpInfo
        objBgImgInfo.BakIp = .strFtpIp
        objBgImgInfo.BakUser = .strFTPUser
        objBgImgInfo.BakPwd = .strFTPPwd
        objBgImgInfo.BakVirtualPath = .strFtpVirtualURL
    End With
    End If
    
    Set GetBgImgInfo = objBgImgInfo
End Function

Public Function GetStudyImgPath(ByRef dcmInfo As TDicomBaseInfo) As String
'��ȡ���ͼ��·��
    Dim strPath As String
    
    strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\" & Format(dcmInfo.strReceiveFullTime, "yyyymmdd") & "\" & dcmInfo.strStudyUID & "\")
    
    If DirExists(strPath) = False Then MkLocalDir strPath
    
    GetStudyImgPath = strPath
End Function


Public Function GetTempImgPath(Optional ByVal blnAutoCreate As Boolean = True) As String
'��ȡ��ʱͼ��·��
    Dim strPath As String
    
    strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\")
    
    If DirExists(strPath) = False And blnAutoCreate Then MkLocalDir strPath
    
    GetTempImgPath = strPath
End Function

Public Function GetCachePath(ByVal strFmtDate As String, Optional ByVal strMark As String = "") As String
'��ȡ����·��
    GetCachePath = FormatFilePath(SysRootPath & "\Apply\TmpAfterImage\" & strFmtDate & "\" & IIf(Len(strMark) <= 0, "", strMark & "\"))
End Function


Public Function GetLineFtpInfo(ByVal strLineDeviceNo As String, ByVal blnMoved As Boolean, ByRef dcmInfo As TDicomBaseInfo, ByRef strErr As String) As TFtpDeviceInf
'��ȡ�µĴ洢�豸��Ϣ������豸�洢��Ϣ�����ڣ�����Ҫ��������

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim curDate As Date
 
On Error GoTo errhandle
    strErr = ""
    
    If mlineFtpInfo.strIID = dcmInfo.strStudyUID Then
        With mlineFtpInfo
            GetLineFtpInfo.strDeviceId = .strDeviceId
            GetLineFtpInfo.strFtpDir = .strFtpDir
            GetLineFtpInfo.strFtpIp = .strFtpIp
            GetLineFtpInfo.strFTPPwd = .strFTPPwd
            GetLineFtpInfo.strFTPUser = .strFTPUser
            GetLineFtpInfo.strFtpVirtualURL = .strFtpVirtualURL
            GetLineFtpInfo.strIID = .strIID
        End With
    Else
        strSQL = "Select D.IP��ַ As Host, D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as λ��,C.��������," & _
            "'/'|| D.FtpĿ¼ ||'/' As Root, Decode(C.��������, Null,'',to_Char(C.��������,'YYYYMMDD') || '/') || C.���UID As URL " & _
            " From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
            " Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
            " And C.���UID= [1]"
        If blnMoved Then strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���洢�豸", dcmInfo.strStudyUID)
    
        blnIsGetNewDevice = False
    
        If rsData.RecordCount <= 0 Then
            blnIsGetNewDevice = True
        Else
            '���ִ�е����˵����ִ��ͼ�����,��Ҫ�жϵ�ǰ���Ĵ洢�豸�Ƿ���Ч�������Ч�������µĴ洢�豸
            If Trim(rsData!��������) = "" Or nvl(rsData!λ��) = "" Then
                blnIsGetNewDevice = True
            Else
                GetLineFtpInfo.strDeviceId = nvl(rsData!λ��)
                GetLineFtpInfo.strFtpIp = nvl(rsData!Host)
                GetLineFtpInfo.strFtpDir = nvl(rsData!Root)
                GetLineFtpInfo.strFTPUser = nvl(rsData!FtpUser)
                GetLineFtpInfo.strFTPPwd = nvl(rsData!FtpPwd)
                GetLineFtpInfo.strFtpVirtualURL = GetLineFtpInfo.strFtpDir & nvl(rsData!Url)
            End If
        End If
    
        If blnIsGetNewDevice Then
            If Val(strLineDeviceNo) <= 0 Then
                strErr = "δ�ҵ�ͼ��洢�豸,��ȷ�϶�Ӧ�洢�豸�ѽ��������á�"
                Exit Function
            End If
     
            strSQL = "Select �豸��,�豸��,'/'|| FtpĿ¼ || '/' As Root,FTP�û���,FTP����,IP��ַ " & _
                        " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ɼ����ߴ洢�豸", strLineDeviceNo)
    
            '����洢�豸ͣ�ã���ֱ���˳�
            If rsTemp.RecordCount <= 0 Then
                strErr = "δ�ҵ����ߴ洢�豸,��ȷ���豸��Ϊ [" & strLineDeviceNo & "] ���豸�Ƿ����á�"
                Exit Function
            End If
    
            GetLineFtpInfo.strDeviceId = strLineDeviceNo
            GetLineFtpInfo.strFtpIp = nvl(rsTemp("IP��ַ"))
            GetLineFtpInfo.strFTPUser = nvl(rsTemp("FTP�û���"))
            GetLineFtpInfo.strFTPPwd = nvl(rsTemp("FTP����"))
            GetLineFtpInfo.strFtpDir = nvl(rsTemp("Root"))
    
            GetLineFtpInfo.strFtpVirtualURL = GetLineFtpInfo.strFtpDir & Format(dcmInfo.strReceiveFullTime, "yyyymmdd") & "/" & dcmInfo.strStudyUID
     
        End If
    End If
    
    mlineFtpInfo = GetLineFtpInfo
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ResetStorageDevice(ByVal lngAdviceId As Long, ByRef objImgInf As clsBgImgInfo, ByVal blnMoved As Boolean) As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    ResetStorageDevice = ""
    
    strSQL = " Select A.���UID,to_Char(A.��������,'YYYYMMDD') As ��������, Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/') ||A.���UID||'/' As URL," & _
            " B.�豸�� as �豸��1, B.�豸�� As �豸��1, B.FTP�û��� As User1,B.FTP���� As Pwd1, B.IP��ַ As Host1, " & _
                    " decode(B.FtpĿ¼, null, '/', '/'||B.FtpĿ¼||'/') As Root1,B.����Ŀ¼ as ����Ŀ¼1,B.����Ŀ¼�û��� as ����Ŀ¼�û���1,B.����Ŀ¼���� as ����Ŀ¼����1 " & _
            " From  Ӱ�����¼ A,Ӱ���豸Ŀ¼ B " & _
            " Where A.ҽ��ID=[1] And Nvl(A.λ��һ, λ�ö�) = B.�豸��(+) "
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��洢", lngAdviceId)
            
    If rsData.RecordCount <= 0 Then
        ResetStorageDevice = "δ�ҵ�����Ӧ��ͼ��洢�豸�����������Ƿ���ȷ��"
        Exit Function
    End If
    
    If nvl(rsData!Host1) <> "" Then
        objImgInf.DeviceNo = Val(nvl(rsData!�豸��1))
        objImgInf.FtpIp = nvl(rsData!Host1)
        objImgInf.FtpUser = nvl(rsData!User1)
        objImgInf.FtpPwd = nvl(rsData!Pwd1)
        objImgInf.FtpVirtualPath = nvl(rsData!Root1) & nvl(rsData!Url)
    End If
    
    objImgInf.AdviceId = lngAdviceId
    objImgInf.StudyUID = nvl(rsData!���UID)
    objImgInf.RecFmtDate = nvl(rsData!��������)
    objImgInf.FilePath = FormatFilePath(SysRootPath & "\Apply\TmpImage\" & nvl(rsData!��������) & "\" & nvl(rsData!���UID) & "\")
 
Exit Function
errhandle:
    ResetStorageDevice = err.Description
End Function

Public Function GetNewStorageDevice(ByVal lngAdviceId As Long, _
    ByVal strStudyUID As String, ByVal strRecFmtDate As String, _
    ByRef objImgInf As clsBgImgInfo) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDeviceNo As String
    Dim strRoot As String
    
    GetNewStorageDevice = ""
    '��ѯҽ������վ�У��������Ӧ�Ĵ洢�豸
    strSQL = "select d.����ֵ " & _
                " from ҽ��ִ�з��� a, ����ҽ������ b, Ӱ��DICOM����� c, Ӱ��DICOM������� d " & _
                " Where a.����ID = b.ִ�в���id And a.ִ�м� = b.ִ�м� And a.����豸 = c.�豸�� " & _
                " and c.������='ͼ�����' and c.����ID=d.����ID and d.��������='�洢�豸' and b.ҽ��id=[1]"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngAdviceId)
    
    If rsTemp.RecordCount <= 0 Then
        GetNewStorageDevice = "δ�ҵ�ͼ��洢�豸,��ȷ�ϵ�ǰ��������豸�Ƿ���Ӱ���豸Ŀ¼�ķ���������������ͼ��洢��"
        Exit Function
    End If
    
    strDeviceNo = nvl(rsTemp!����ֵ)


    strSQL = "Select �豸��,�豸��,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼ || '/') As Root,FTP�û���,FTP����,IP��ַ " & _
                " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strDeviceNo)
    
    '����洢�豸ͣ�ã���ֱ���˳�
    If rsTemp.RecordCount <= 0 Then
        GetNewStorageDevice = "δ�ҵ��洢�豸,��ȷ���豸��Ϊ [" & strDeviceNo & "] ���豸�Ƿ����á�"
        Exit Function
    End If
    
    
    strRoot = nvl(rsTemp("Root"))
    
    objImgInf.DeviceNo = Val(strDeviceNo)
    objImgInf.AdviceId = lngAdviceId
    objImgInf.StudyUID = strStudyUID
    objImgInf.RecFmtDate = strRecFmtDate
    
    objImgInf.FtpIp = nvl(rsTemp("IP��ַ"))
    objImgInf.FtpUser = nvl(rsTemp("FTP�û���"))
    objImgInf.FtpPwd = nvl(rsTemp("FTP����"))
    objImgInf.FtpVirtualPath = IIf(strRoot = "/", "//", strRoot) & strRecFmtDate & "/" & strStudyUID & "/"
    
    
    objImgInf.FilePath = FormatFilePath(SysRootPath & "\Apply\TmpImage\" & strRecFmtDate & "\" & strStudyUID & "\")

End Function

Public Function GetBackFtpInfo(ByVal strBackDeviceNo As String, ByRef dcmInfo As TDicomBaseInfo, ByRef strErr As String) As TFtpDeviceInf
'��ȡ�µĴ洢�豸��Ϣ������豸�洢��Ϣ�����ڣ�����Ҫ��������

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim curDate As Date
 
On Error GoTo errhandle
    strErr = ""
    
    If mbackFtpInfo.strIID = dcmInfo.strStudyUID Then
        With mbackFtpInfo
            GetBackFtpInfo.strDeviceId = .strDeviceId
            GetBackFtpInfo.strFtpDir = .strFtpDir
            GetBackFtpInfo.strFtpIp = .strFtpIp
            GetBackFtpInfo.strFTPPwd = .strFTPPwd
            GetBackFtpInfo.strFTPUser = .strFTPUser
            GetBackFtpInfo.strFtpVirtualURL = .strFtpVirtualURL
            GetBackFtpInfo.strIID = .strIID
        End With
    Else
        If Len(strBackDeviceNo) <= 0 Then Exit Function
    
        strSQL = "Select �豸��,�豸��,'/'|| FtpĿ¼ || '/' As Root,FTP�û���,FTP����,IP��ַ " & _
                    " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ɼ����ݴ洢�豸", strBackDeviceNo)

        '����洢�豸ͣ�ã���ֱ���˳�
        If rsTemp.RecordCount <= 0 Then
            strErr = "δ�ҵ����ݴ洢�豸,��ȷ���豸��Ϊ [" & strBackDeviceNo & "] ���豸�Ƿ����á�"
            Exit Function
        End If

        GetBackFtpInfo.strDeviceId = strBackDeviceNo
        GetBackFtpInfo.strFtpIp = nvl(rsTemp("IP��ַ"))
        GetBackFtpInfo.strFTPUser = nvl(rsTemp("FTP�û���"))
        GetBackFtpInfo.strFTPPwd = nvl(rsTemp("FTP����"))
        GetBackFtpInfo.strFtpDir = nvl(rsTemp("Root"))

        GetBackFtpInfo.strFtpVirtualURL = GetBackFtpInfo.strFtpDir & Format(dcmInfo.strReceiveFullTime, "yyyymmdd") & "/" & dcmInfo.strStudyUID
      
    End If
    
    mbackFtpInfo = GetBackFtpInfo
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDicomAge(ByVal dtBirth As String, Optional ByVal strAge As String = "") As String
'��ȡDicom�����ʽ
    Dim dtStart As Date
    Dim lngDays As Long
    Dim lngAge As Long
    
    GetDicomAge = ""
    
    If Len(dtBirth) > 0 Then
        dtStart = CDate(Format(dtBirth, "yyyy-mm-dd"))
        lngDays = DateDiff("d", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
        
        'ת��Ϊ��,��,��,��
        Select Case True
            Case lngDays > 365 * 3 '3��
                '��
                GetDicomAge = DateDiff("yyyy", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "Y"
            Case lngDays > 30 * 3 '3��
                '��
                GetDicomAge = DateDiff("m", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "M"
            Case lngDays > 7 * 4 'һ��
                '��
                GetDicomAge = DateDiff("ww", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "W"
            Case Else
                '��
                GetDicomAge = DateDiff("d", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "D"
        End Select
        
        Exit Function
    End If
    
    If Len(strAge) > 0 Then
        '����¼�������ת��Ϊdicom��ʽ��������ʽ
        lngAge = Val(strAge)
        
        Select Case True
            Case (InStr(strAge, "��") > 0), (InStr(UCase(strAge), "Y") > 0):
                GetDicomAge = Format(lngAge, "000") & "Y"
            Case (InStr(strAge, "��") > 0), (InStr(UCase(strAge), "M") > 0):
                GetDicomAge = Format(lngAge, "000") & "M"
            Case (InStr(strAge, "��") > 0), (InStr(UCase(strAge), "W") > 0):
                GetDicomAge = Format(lngAge, "000") & "W"
            Case Else
                GetDicomAge = Format(lngAge, "000") & "D"
        End Select
    End If
    
End Function


Public Function GetDicomBaseInfoEx(ByVal lngAdviceId As Long, dcmImg As DicomImage, Optional ByRef strDeviceNo As String) As TDicomBaseInfo
    Dim objValue As Variant
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    
    If dcmImg Is Nothing Then Exit Function
     
    If mlastDicomInfo.lngAdviceId = lngAdviceId Then
        With mlastDicomInfo
            GetDicomBaseInfoEx.lngAdviceId = .lngAdviceId
            GetDicomBaseInfoEx.lngSendNo = .lngSendNo
            GetDicomBaseInfoEx.lngID = .lngID
            GetDicomBaseInfoEx.strAge = .strAge
            GetDicomBaseInfoEx.strBirthDate = .strBirthDate
            
            GetDicomBaseInfoEx.strInstitution = .strInstitution
            GetDicomBaseInfoEx.strModality = .strModality
            GetDicomBaseInfoEx.strName = .strName
            GetDicomBaseInfoEx.strSex = .strSex
            GetDicomBaseInfoEx.strReceiveFullTime = IIf(Len(.strReceiveFullTime) > 0, .strReceiveFullTime, zlDatabase.Currentdate)
            
            GetDicomBaseInfoEx.strStudyUID = .strStudyUID
            GetDicomBaseInfoEx.strSeriesUID = .strSeriesUID
            GetDicomBaseInfoEx.strDeviceNo = .strDeviceNo
             
        End With
    Else
        strSQL = "select b.����ID,a.���ͺ�,a.Ӱ�����,a.����豸,nvl(a.λ��һ, a.λ�ö�) as �洢λ��, a.����,a.�Ա�,a.��������,a.����,a.���UID,a.�������� " & _
                " From Ӱ�����¼ a, ����ҽ����¼ b " & _
                " Where a.ҽ��ID=b.Id and a.ҽ��ID=[1]"
                
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����Ϣ", lngAdviceId)
        
        If rsData.RecordCount <= 0 Then Exit Function
    
        GetDicomBaseInfoEx.lngAdviceId = lngAdviceId
        GetDicomBaseInfoEx.lngSendNo = Val(nvl(rsData!���ͺ�))
        GetDicomBaseInfoEx.lngID = Val(nvl(rsData!����ID))
        GetDicomBaseInfoEx.strAge = GetDicomAge(nvl(rsData!��������), nvl(rsData!����))
        GetDicomBaseInfoEx.strBirthDate = Format(nvl(rsData!��������), "yyyymmdd")
        GetDicomBaseInfoEx.strInstitution = RegInstitution
        GetDicomBaseInfoEx.strModality = nvl(rsData!Ӱ�����)
        GetDicomBaseInfoEx.strName = nvl(rsData!����)
        GetDicomBaseInfoEx.strSex = Decode(nvl(rsData!�Ա�), "��", "M", "Ů", "F", "O")
        GetDicomBaseInfoEx.strReceiveFullTime = nvl(rsData!��������, zlDatabase.Currentdate) ', "yyyymmdd")
        GetDicomBaseInfoEx.strStudyUID = nvl(rsData!���UID)
        GetDicomBaseInfoEx.strDeviceNo = nvl(rsData!�洢λ��)
    End If
    
    objValue = dcmImg.Attributes(&H10, &H20).value      '����ID
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.lngID = Val(objValue)
     
    objValue = dcmImg.Attributes(&H10, &H1010).value      '��������
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strAge = objValue
     
    objValue = dcmImg.Attributes(&H10, &H30).value      '���˳�������
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strBirthDate = dcmImg.DateOfBirthAsDate
     
    objValue = dcmImg.Attributes(&H8, &H80).value      '��λ����
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strInstitution = objValue
     
    objValue = dcmImg.Attributes(&H8, &H60).value      'Ӱ�����
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strInstitution = objValue
      
    If Len(dcmImg.StudyUID) > 0 Then GetDicomBaseInfoEx.strStudyUID = dcmImg.StudyUID     '���UID
    If Len(dcmImg.SeriesUID) > 0 Then GetDicomBaseInfoEx.strSeriesUID = dcmImg.SeriesUID  '����UID
    If Len(dcmImg.InstanceUID) > 0 Then GetDicomBaseInfoEx.strInstanceUID = dcmImg.InstanceUID    'ʵ��UID
    
    GetDicomBaseInfoEx.lngSeriesNo = Val(dcmImg.Attributes(&H20, &H11).value)  '���к�
    GetDicomBaseInfoEx.lngImgNo = Val(dcmImg.Attributes(&H20, &H13).value)     'ͼ���
    
    mlastDicomInfo = GetDicomBaseInfoEx
End Function


Public Function GetDicomBaseInfo(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As TDicomBaseInfo
'��ȡDicom������Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    If mlastDicomInfo.lngAdviceId = lngAdviceId Then
        With mlastDicomInfo
            GetDicomBaseInfo.lngAdviceId = .lngAdviceId
            GetDicomBaseInfo.lngSendNo = .lngSendNo
            GetDicomBaseInfo.lngID = .lngID
            GetDicomBaseInfo.strAge = .strAge
            GetDicomBaseInfo.strBirthDate = .strBirthDate
            
            GetDicomBaseInfo.strInstitution = .strInstitution
            GetDicomBaseInfo.strModality = .strModality
            GetDicomBaseInfo.strName = .strName
            GetDicomBaseInfo.strSex = .strSex
            GetDicomBaseInfo.strReceiveFullTime = .strReceiveFullTime
            
            GetDicomBaseInfo.strStudyUID = .strStudyUID
            GetDicomBaseInfo.strSeriesUID = .strSeriesUID
            GetDicomBaseInfo.strInstanceUID = CreateUID
            
            GetDicomBaseInfo.lngSeriesNo = .lngSeriesNo
            
            GetDicomBaseInfo.lngImgNo = .lngImgNo + 1
        End With
    Else
        strSQL = "select b.����ID,a.���ͺ�,a.Ӱ�����,a.����豸,a.����,a.�Ա�,a.��������,a.����,a.���UID,a.�������� " & _
                " From Ӱ�����¼ a, ����ҽ����¼ b " & _
                " Where a.ҽ��ID=b.Id and a.ҽ��ID=[1]"
                
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����Ϣ", lngAdviceId)
        
        If rsData.RecordCount <= 0 Then Exit Function
    
        GetDicomBaseInfo.lngAdviceId = lngAdviceId
        GetDicomBaseInfo.lngSendNo = Val(nvl(rsData!���ͺ�))
        GetDicomBaseInfo.lngID = Val(nvl(rsData!����ID))
        GetDicomBaseInfo.strAge = GetDicomAge(nvl(rsData!��������), nvl(rsData!����))
        GetDicomBaseInfo.strBirthDate = Format(nvl(rsData!��������), "yyyymmdd")
        GetDicomBaseInfo.strInstanceUID = CreateUID
        GetDicomBaseInfo.strInstitution = RegInstitution
        GetDicomBaseInfo.strModality = nvl(rsData!Ӱ�����)
        GetDicomBaseInfo.strName = nvl(rsData!����)
        GetDicomBaseInfo.strSex = Decode(nvl(rsData!�Ա�), "��", "M", "Ů", "F", "O")
        GetDicomBaseInfo.strReceiveFullTime = nvl(rsData!��������, zlDatabase.Currentdate) ', "yyyymmdd")
        
        GetDicomBaseInfo.strStudyUID = nvl(rsData!���UID)
        GetDicomBaseInfo.lngSeriesNo = 1
        GetDicomBaseInfo.lngImgNo = 1
        
        If Len(GetDicomBaseInfo.strStudyUID) > 0 Then   'lngImgNo=0��ʾ��һ�βɼ�ͼ��
            '��ȡ����UID��ͼ���
            strSQL = "Select ����UID,���к� From Ӱ�������� Where ���UID=[1] and �������� is Null order by ���к�"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ϣ", GetDicomBaseInfo.strStudyUID)
            
            If rsData.RecordCount > 0 Then
                GetDicomBaseInfo.strSeriesUID = nvl(rsData!����UID)
                GetDicomBaseInfo.lngSeriesNo = Val(nvl(rsData!���к�))
                
                strSQL = "select max(nvl(ͼ���, 0)) as ͼ��� From Ӱ����ͼ�� Where ����UID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ���", GetDicomBaseInfo.strSeriesUID)
                
                If rsData.RecordCount > 0 Then
                    GetDicomBaseInfo.lngImgNo = Val(nvl(rsData!ͼ���)) + 1
                End If
                
            Else
                GetDicomBaseInfo.strSeriesUID = CreateUID
            End If
        Else
            GetDicomBaseInfo.strStudyUID = CreateUID     '��ȡ�µļ��UID
            GetDicomBaseInfo.strSeriesUID = CreateUID    '��ȡ�µ�����UID
        End If
    End If
    
    mlastDicomInfo = GetDicomBaseInfo
End Function


Public Sub WriteDicomPara(img As DicomImage, dicomInfo As TDicomBaseInfo, _
    Optional blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'���ܣ��������ͼ����дDICOM�ļ�ͷ��Ϣ
'������img���������DICOM�ļ�,lngAdviceID����ҽ��ID
'���أ��ޣ�ֱ���ļ�ͷ��Ϣд��img���ļ�ͷ
'------------------------------------------------
    Dim curDate As Date

    curDate = zlDatabase.Currentdate
    
    If blnIsAfterCapture Then
        img.Attributes.Add &H10, &H10, ""                           'Name ����
        img.Attributes.Add &H10, &H20, ""                           'Patient ID ����ID
        img.Attributes.Add &H10, &H30, ""                           'BirthDate ����
        img.Attributes.Add &H10, &H40, ""                           'Sex �Ա�
        img.Attributes.Add &H10, &H1010, ""                         'Age ����
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
        img.Attributes.Add &H20, &H10, ""                           'Study ID ���ID
        img.Attributes.Add &H8, &H60, dicomInfo.strModality         'Modality Ӱ�����
        img.Attributes.Add &H20, &H11, "1"                          'Series Number ���к�
        img.Attributes.Add &H20, &H13, "1"                          'ImageNumber ͼ���
    Else
        img.Attributes.Add &H10, &H10, dicomInfo.strName            'Name ����
        img.Attributes.Add &H10, &H20, dicomInfo.lngID              'Patient ID ����ID
        img.Attributes.Add &H10, &H30, dicomInfo.strBirthDate       'BirthDate ����
        img.Attributes.Add &H10, &H40, dicomInfo.strSex             'Sex �Ա�
        img.Attributes.Add &H10, &H1010, dicomInfo.strAge           'Age ����
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
        img.Attributes.Add &H8, &H60, dicomInfo.strModality         'Modality Ӱ�����
        
        img.StudyUID = dicomInfo.strStudyUID                        ' &H20, &H10 Study ID ���ID
        img.SeriesUID = dicomInfo.strSeriesUID                      ' ����UID
        img.InstanceUID = dicomInfo.strInstanceUID                  'ͼ��ʵ��UID
        
        img.Attributes.Add &H20, &H11, dicomInfo.lngSeriesNo        'Series Number ���к�
        img.Attributes.Add &H20, &H13, dicomInfo.lngImgNo           'ImageNumber ͼ���
    End If
    
    img.Attributes.Add &H8, &H8, ""                                 'ImageType  ��
    img.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"       'SOP Class  UID�����β�׽
    img.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date �������
    img.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date ��������
    img.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date �ɼ�����
    img.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   ͼ������
    img.Attributes.Add &H8, &H30, Format(curDate, "HH24:MI:SS")     'Study Time   ���ʱ��
    img.Attributes.Add &H8, &H31, Format(curDate, "HH24:MI:SS")     'Series Time  ����ʱ��
    img.Attributes.Add &H8, &H32, Format(curDate, "HH24:MI:SS")     'Acquisition Time  �ɼ�ʱ��
    img.Attributes.Add &H8, &H33, Format(curDate, "HH24:MI:SS")     'Image Time  ͼ��ʱ��
    img.Attributes.Add &H8, &H50, ""                            'Accession Number ��
    img.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer ����
    img.Attributes.Add &H8, &H80, RegInstitution               'Institution Name ��λ����
    img.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name ��
    img.Attributes.Add &H8, &H1030, ""                          'Study Description ������� ��

    img.Attributes.Add &H20, &H20, ""                           'Orientation ��
    
End Sub

Public Sub SaveImageInfo(ByRef dcmInfo As TDicomBaseInfo, ByRef ftpInfo As TFtpDeviceInf)
'����ɼ�ͼ��
    Dim arySql() As String
    Dim strSQL As String
    Dim blnInTrans As Boolean
    
    Dim rsData As ADODB.Recordset
    Dim blnHasStudy As Boolean
    Dim blnHasSeries As Boolean
    
    Dim i As Long
    
On Error GoTo errhandle

    ReDim arySql(0)
    
    strSQL = "select 1 from Ӱ�����¼ where ���UID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����Ϣ", dcmInfo.strStudyUID)
    
    blnHasStudy = IIf(rsData.RecordCount <= 0, False, True)
    If blnHasStudy Then
        '�ж��Ƿ��ж�Ӧ��������Ϣ
        strSQL = "select 1 from Ӱ�������� where ����UID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����Ϣ", dcmInfo.strSeriesUID)
        
        blnHasSeries = IIf(rsData.RecordCount <= 0, False, True)
    Else
        blnHasSeries = False
    End If
    
    If blnHasStudy = False Then
        '�״βɼ�ͼ��,��Ҫд��ɼ�������Ϣ�ʹ洢�豸��Ϣ...
        strSQL = "ZL_Ӱ�����¼_SET(" & dcmInfo.lngAdviceId & "," & dcmInfo.lngSendNo & ",'" & _
                                        dcmInfo.strStudyUID & "',null," & _
                                        "to_Date('" & Format(dcmInfo.strReceiveFullTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                                        ftpInfo.strDeviceId & "')"
                                        
        ReDim Preserve arySql(UBound(arySql) + 1)
        arySql(UBound(arySql)) = strSQL
        
        dcmInfo.lngImgNo = 1
    End If
    
    If blnHasSeries = False Then
        strSQL = "ZL_Ӱ������_INSERT('" & dcmInfo.strStudyUID & "','" & dcmInfo.strSeriesUID & "','" & dcmInfo.strSeriesDes & "',0)"
        
        ReDim Preserve arySql(UBound(arySql) + 1)
        arySql(UBound(arySql)) = strSQL
    End If
    
    If dcmInfo.lngMediaTag = 0 Then
        strSQL = "ZL_Ӱ��ͼ��_INSERT('" & dcmInfo.strInstanceUID & "','" & dcmInfo.strSeriesUID & "',NULL,0, null, sysdate)"
    Else
        strSQL = "ZL_Ӱ��ͼ��_INSERT('" & dcmInfo.strInstanceUID & "','" & dcmInfo.strSeriesUID & "',Null,0" & _
        ",null,sysdate,null,null,null,null,null,null,null,null,null," & dcmInfo.lngMediaTag & ",'" & dcmInfo.strMediaEncode & "'," & dcmInfo.lngMediaLen & ")"
    End If
    
    ReDim Preserve arySql(UBound(arySql) + 1)
    arySql(UBound(arySql)) = strSQL
        
    gcnOracle.BeginTrans        '----------����ý�壬ͼ����Ƶ����Ƶ
    blnInTrans = True
    
    For i = 1 To UBound(arySql)
        Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), "����ִ�вɼ�ý�屣��")
    Next i
    
    gcnOracle.CommitTrans
Exit Sub
errhandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function IsValidFile(ByVal strFile As String, Optional ByVal lngSize As Long = 0)
    Dim objFileSystem As New FileSystemObject
    Dim lngFileSize As Long
    
On Error GoTo errhandle
    IsValidFile = False
    
    If Trim(Dir(strFile, 7)) = "" Then Exit Function
    
    lngFileSize = lngSize
    If lngFileSize <= 0 Then lngFileSize = 1000
    
    If objFileSystem.GetFile(strFile).Size < lngFileSize Then Exit Function
    
    IsValidFile = True
    
    Set objFileSystem = Nothing
Exit Function
errhandle:
    IsValidFile = False
    Set objFileSystem = Nothing
End Function

Private Function IsFileLocked(ByVal strFileName As String) As Boolean
   Dim iFn As Integer
   Dim blnRetVal As Boolean
   
On Error GoTo E_HandleFA
    blnRetVal = True
    
    If (Len(Dir$(strFileName, 7)) > 0) Then
       iFn = FreeFile
       
       Open strFileName For Binary Lock Read Write As #iFn
       Close iFn
       
       blnRetVal = False
       
       blnRetVal = IsFileOpen(strFileName)
    Else
        '�ļ������ڣ��򷵻�δ����״̬
        blnRetVal = False
    End If
   
E_HandleFA:
   IsFileLocked = blnRetVal
End Function


Private Function IsFileOpen(ByVal pFile As String) As Boolean
    Dim ret As Long
    
    ret = CreateFile(pFile, GENERIC_READ Or GENERIC_WRITE, 0&, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    
    IsFileOpen = (ret = INVALID_HANDLE_VALUE)
    
    CloseHandle ret
End Function

Private Function WaitReadDcm(dImgs As DicomImages, ByVal strFile As String, _
    ByRef strError As String) As DicomImage
On Error Resume Next
    Dim i As Long
    Dim blnUseUrl As Boolean
    
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    Set WaitReadDcm = Nothing
    
    While True
        err.Clear
        dImgs.Clear
        
        If blnUseUrl Then
            'readurl��֧�ֿո�
            Set WaitReadDcm = dImgs.ReadURL(strFile)
        Else
            Set WaitReadDcm = dImgs.ReadFile(strFile)
        End If
        
        If err.Description = "" Then Exit Function
        
        i = i + 1
         
        If i > 100 Then
            strError = err.Description
            Exit Function
        End If
        
        Call Sleep(10)
    Wend
End Function

Private Function GetFileSize(ByVal strFile As String) As Long
'��ȡ�ļ���С
On Error GoTo errhandle
    GetFileSize = FileLen(strFile)
Exit Function
errhandle:
    GetFileSize = 0
End Function

Public Function ReadDicomFile(ByVal strFile As String, ByRef strError As String, _
    Optional ByVal blnIsDcmFormat As Boolean = False) As DicomImage
On Error Resume Next
    Dim dImgs As New DicomImages
        
    Dim curImage As DicomImage
    Dim blnUseUrl As Boolean
    Dim strFileTime As String
    Dim strCopyFileName As String
    Dim lngSize As Long
    
    strError = ""
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    
    '����ռλ
    If strFile = "NULL" Then
        Set curImage = dImgs.AddNew
    
        dImgs.Clear
        Set dImgs = Nothing
        
        Set ReadDicomFile = curImage
        
        Exit Function
    End If
    
    If blnUseUrl Then
        'readurl��֧�ֿո�
        Set curImage = dImgs.ReadURL(strFile)
    Else
        Set curImage = dImgs.ReadFile(strFile)
    End If
    
    If err.Number = 0 Then
        If Not curImage Is Nothing Then
            If Len(curImage.InstanceUID) > 0 Then
                dImgs.Clear
                Set dImgs = Nothing
                
                Set ReadDicomFile = curImage
                Exit Function
            End If
        End If
    End If
    

    '2098����һ�����ļ�����dicom�ļ�����һ���Ǵ��ڹ�����ʴ���
    If InStr(err.Description, "sharing violation") > 0 Then
        
        lngSize = GetFileSize(strFile)
        
        strFileTime = Format(Now, "MMDD") & GetTickCount
        strCopyFileName = strFile & "_copy_vdat_" & strFileTime
        
        Call FileCopy(strFile, strCopyFileName)
        
        err.Clear
        
        If IsValidFile(strCopyFileName, lngSize) = False Then
            '�ļ�����ʧ�ܣ��������¸���
            If WaitCopy(strFile, strCopyFileName, strError, lngSize) = False Then
                '�ļ�����ʧ��
                dImgs.Clear
                Set dImgs = Nothing
                 
                Set ReadDicomFile = Nothing
                Exit Function
            End If
        End If
    
        If blnUseUrl Then
            'readurl��֧�ֿո�
            Set curImage = dImgs.ReadURL(strCopyFileName)
        Else
            Set curImage = dImgs.ReadFile(strCopyFileName)
        End If
        
        If curImage Is Nothing Or err.Number <> 0 Then
            err.Clear
            Set curImage = WaitReadDcm(dImgs, strCopyFileName, strError)
        End If
        
        If err.Number = 0 Then
            Call Kill(strCopyFileName)
            'ʹ��ReadFile��ʽ��ȡ���ļ�����ɾ��ʱ�����ܻ�����쳣
            err.Clear
        Else
            Call Kill(strCopyFileName)
        End If
    Else
        If blnIsDcmFormat = False Then
            err.Clear
            Set curImage = dImgs.AddNew
            Call curImage.FileImport(strFile, "JPG")
            
            If err.Number <> 0 Then
                err.Clear
                'not a JPG file
                Call curImage.FileImport(strFile, "BMP")
            End If
            
            If err.Number <> 0 Then
                '�ļ����쳣ʱ��ɾ����ӵ�item��
                Call dImgs.Remove(dImgs.Count)
            End If
        Else
            'ָ����ȡdicom�ļ��ǣ���Ҫ���������ݴ���
            If err.Number <> 0 Or curImage Is Nothing Then
                Set curImage = WaitReadDcm(dImgs, strFile, strError)
                
                If Not curImage Is Nothing Then err.Clear
            End If
        End If
    End If
    
    dImgs.Clear
    
    Set dImgs = Nothing
    Set ReadDicomFile = Nothing
    
    If err.Number = 0 Then
        If curImage Is Nothing Then
            strError = "�ļ���ʽ��ȡ����"
            Exit Function
        End If
        
        If Len(curImage.InstanceUID) <= 0 Then
            strError = "DICOM��ʽ�ļ�����,δ��ȡ��ʵ��UID"
            Exit Function
        End If
        
        Set ReadDicomFile = curImage
    Else
        strError = err.Description
    End If
    
End Function

Private Function WaitCopy(ByVal strSourceFile As String, ByVal strTargetFile As String, _
    ByRef strError As String, Optional ByVal lngSize As Long = 0) As Boolean
    Dim i As Long
On Error Resume Next
    WaitCopy = False
    
    i = 0
    While True
    
        If IsFileLocked(strTargetFile) = False Then
            Call FileCopy(strSourceFile, strTargetFile)
        
            If IsValidFile(strTargetFile, lngSize) <> "" Then
                WaitCopy = True
                Exit Function
            End If
        End If
        
        i = i + 1
        
        If i > 300 Then
            strError = err.Description
            Exit Function
        End If
        Sleep 10
    Wend
    
End Function


Public Function HasProcess(ByVal strAppTitle As String)
    Dim lngDeskTopHandle As Long
    Dim lngHand As Long
    Dim strName As String * 255
    Dim strCurAppName As String
    
On Error GoTo errhandle
    lngDeskTopHandle = GetDesktopWindow()
    lngHand = GetWindow(lngDeskTopHandle, GW_CHILD)
    
    Do While lngHand <> 0
       GetWindowText lngHand, strName, Len(strName)
       lngHand = GetWindow(lngHand, GW_HWNDNEXT)
       If Left$(strName, 1) <> vbNullChar Then
          strCurAppName = Left$(strName, InStr(1, strName, vbNullChar) - 1)
          If UCase(strCurAppName) = strAppTitle Then
              HasProcess = True
              Exit Function
          End If
       End If
    Loop
     
    HasProcess = False
Exit Function
errhandle:
    HasProcess = False
End Function


Public Function FormatFilePath(ByVal strFilePath As String) As String
'��ʽ���ļ�·��
    FormatFilePath = Replace(strFilePath, "\\", "\")
End Function


Public Function GetImgCmdPath(Optional ByVal blnIsFailed = False) As String
    Dim strPath As String
    
    If blnIsFailed Then
        strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\TransCmd\Failed\")
    Else
        strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\TransCmd\")
    End If
    
    If DirExists(strPath) = False Then
        Call MkLocalDir(strPath)
    End If
    
    GetImgCmdPath = strPath
End Function


Public Function GetImgCmdFile(objImgInfo As clsBgImgInfo) As String
    Dim strPath As String
    
    strPath = GetImgCmdPath
    
    GetImgCmdFile = FormatFilePath(strPath & objImgInfo.Key)
End Function

Public Function GetImgCmdFailed(objImgInfo As clsBgImgInfo) As String
    Dim strPath As String
    
    strPath = GetImgCmdPath(True)
    
    GetImgCmdFailed = FormatFilePath(strPath & objImgInfo.Key)
End Function


Public Sub SetFileHide(ByVal strFile As String)
    Dim dwAtrr   As Long
    
    '�Ȼ�ȡԭ�����ļ�����
    dwAtrr = GetFileAttributes(strFile)
    '������������
    dwAtrr = dwAtrr Or FILE_ATTRIBUTE_HIDDEN
    'ȥ����������
    'dwAtrr = dwAtrr And Not FILE_ATTRIBUTE_HIDDEN
    '�����µ��ļ�����
    Call SetFileAttributes(strFile, dwAtrr)
End Sub


Private Sub HideFile(ByVal strFile As String)
    Dim oFileSystem As New FileSystemObject
    Dim oFile As File
    
    Set oFile = oFileSystem.GetFile(strFile)
    
    oFile.Attributes = 2
    
    Set oFile = Nothing
    Set oFileSystem = Nothing
End Sub


Public Function TransCmd(imgInfo As clsBgImgInfo, ByVal strCmdFile As String, ByRef strError As String) As Boolean
    Dim blnStartState As Boolean
On Error GoTo errhandle
    TransCmd = False
    strError = ""
    
    '�ָ�״̬����
    imgInfo.Redo = 0
    imgInfo.ErrorInfo = ""
    imgInfo.StartTime = Now
    imgInfo.EndTime = 0
    
    TransCmd = CreateCmdFileEx(imgInfo, strCmdFile, strError)
    If TransCmd = False Then Exit Function
    
    imgInfo.LoadState = lsSent
 
    TransCmd = True
Exit Function
errhandle:
    strError = err.Description
End Function

Private Function CreateCmdFile(imgInfo As clsBgImgInfo, ByVal strCmdFile As String, ByRef strError As String) As Boolean
'�������ݽ��������ļ�
'���ݽ���ͨ��Ini�ļ����У�ֻ�����غ���ļ���������ȡ���ļ�����key�������ļ�Ŀ¼ΪTransCmd
    Dim strFile As String
    Dim objIni As New clsIniFile
    
On Error GoTo errhandle
    CreateCmdFile = False
    
    strFile = strCmdFile
    
    If Trim(Dir(strFile, 7)) <> "" Then
        '�����Ѿ�������ֱ���˳�
        CreateCmdFile = True
        Exit Function
    End If
    
    Call objIni.SetIniFile(strFile)
    
    With imgInfo
        objIni.WriteValue "BASEINFO", "KEY", .Key
        objIni.WriteValue "BASEINFO", "FILENAME", .Filename
        objIni.WriteValue "BASEINFO", "FILEPATH", .FilePath
        objIni.WriteValue "BASEINFO", "FORMAT", .Format
        objIni.WriteValue "BASEINFO", "PATIENTNAME", .PatientName
        objIni.WriteValue "BASEINFO", "ADVICEID", .AdviceId
        objIni.WriteValue "BASEINFO", "ADVICEDES", .AdviceDes
        
        objIni.WriteValue "FTPINFO", "FTPIP", .FtpIp
        objIni.WriteValue "FTPINFO", "FTPPORT", .FtpPort
        objIni.WriteValue "FTPINFO", "FTPUSER", .FtpUser
        objIni.WriteValue "FTPINFO", "FTPPWD", .FtpPwd
        objIni.WriteValue "FTPINFO", "FTPVIRTUALPATH", .FtpVirtualPath
        objIni.WriteValue "FTPINFO", "FTPSHDIR", .FtpShareDir
        objIni.WriteValue "FTPINFO", "FTPSHUSER", .FtpShareUser
        objIni.WriteValue "FTPINFO", "FTPSHPWD", .FtpSharePwd
        objIni.WriteValue "FTPINFO", "FTPFILE", .FtpFile
         
        objIni.WriteValue "OTHERINFO", "IMGCOMMAND", .ImgCommand
        objIni.WriteValue "OTHERINFO", "STARTTIME", Now
        objIni.WriteValue "OTHERINFO", "ENDTIME", 0
        objIni.WriteValue "OTHERINFO", "REDO", 0
        
        objIni.WriteValue "OTHERINFO", "ISCOMPRESS", CStr(.IsCompress)
        objIni.WriteValue "OTHERINFO", "JPGCONVERT", CStr(.JpgConvert)
    End With
    
    Set objIni = Nothing
    
    '�����ļ�
'    HideFile strFile   '�÷�������ɽ��̴�����
    Call SetFileHide(strFile)
     
    CreateCmdFile = True
Exit Function
errhandle:
    CreateCmdFile = False
    strError = err.Description
End Function


Private Function CreateCmdFileEx(imgInfo As clsBgImgInfo, ByVal strCmdFile As String, ByRef strError As String) As Boolean
'�������ݽ��������ļ�
'���ݽ���ͨ��Ini�ļ����У�ֻ�����غ���ļ���������ȡ���ļ�����key�������ļ�Ŀ¼ΪTransCmd
    Dim strFile As String
'    Dim objIni As New clsIniFile
    Dim strIniContext As String
    
On Error GoTo errhandle
    CreateCmdFileEx = False
    
    strFile = strCmdFile
    
    If Trim(Dir(strFile, 7)) <> "" Then
        '�����Ѿ�������ֱ���˳�
        CreateCmdFileEx = True
        Exit Function
    End If
    
'    Call objIni.SetIniFile(strFile)
    
    With imgInfo
        strIniContext = "[BASEINFO]" & vbCrLf & _
                                    "KEY=" & .Key & vbCrLf & _
                                    "FILENAME=" & .Filename & vbCrLf & _
                                    "FILEPATH=" & .FilePath & vbCrLf & _
                                    "FORMAT=" & .Format & vbCrLf & _
                                    "PATIENTNAME=" & .PatientName & vbCrLf & _
                                    "ADVICEID=" & .AdviceId & vbCrLf & _
                                    "ADVICEDES=" & .AdviceDes & vbCrLf & _
                                    "[FTPINFO]" & vbCrLf & _
                                    "FTPIP=" & .FtpIp & vbCrLf & _
                                    "FTPPORT=" & .FtpPort & vbCrLf & _
                                    "FTPUSER=" & .FtpUser & vbCrLf & _
                                    "FTPPWD=" & .FtpPwd & vbCrLf & _
                                    "FTPVIRTUALPATH=" & .FtpVirtualPath & vbCrLf & _
                                    "FTPSHDIR=" & .FtpShareDir & vbCrLf & _
                                    "FTPSHUSER=" & .FtpShareUser & vbCrLf & _
                                    "FTPSHPWD=" & .FtpSharePwd & vbCrLf & _
                                    "FTPFILE=" & .FtpFile & vbCrLf & _
                                    "[OTHERINFO]" & vbCrLf & _
                                    "IMGCOMMAND=" & .ImgCommand & vbCrLf & _
                                    "STARTTIME=" & Now & vbCrLf & _
                                    "ENDTIME=0" & vbCrLf & _
                                    "REDO=0" & vbCrLf & _
                                    "ISCOMPRESS=" & CStr(.IsCompress) & vbCrLf & _
                                    "JPGCONVERT=" & CStr(.JpgConvert)
    End With
    
    Call WritTextFile(strFile, strIniContext)
'    Set objIni = Nothing
    
    '�����ļ�
'    HideFile strFile   '�÷�������ɽ��̴�����
    Call SetFileHide(strFile)
     
    CreateCmdFileEx = True
Exit Function
errhandle:
    CreateCmdFileEx = False
    strError = err.Description
End Function

Public Sub DrawErrorText(objImg As DicomImage, ByVal strError As String)
'���ƴ����ı�
    Dim i As Long
    Dim objLabInfo As DicomLabel
    
    Set objLabInfo = Nothing
    
    For i = 1 To objImg.Labels.Count
        If objImg.Labels(i).tag = IMG_LAB_ERRORINFO_TAG Then
            Set objLabInfo = objImg.Labels(i)
        End If
    Next
    
    If objLabInfo Is Nothing Then
        Set objLabInfo = objImg.Labels.AddNew
        objLabInfo.tag = IMG_LAB_ERRORINFO_TAG
    End If
    
    'Text*********************************************
    objLabInfo.LabelType = doLabelText
    objLabInfo.Margin = 0
    objLabInfo.FontSize = 10
    objLabInfo.AutoSize = True
    
    objLabInfo.Text = "��" & strError
    objLabInfo.ForeColour = vbYellow ' vbRed
'    objLabInfo.BackColour = vbYellow
    
    objLabInfo.Transparent = True
    objLabInfo.ScaleWithCell = False
     
    objLabInfo.Left = 40
    objLabInfo.Top = 2
'    objLabInfo.Width = 1000
'    objLabInfo.Height = 1000
    
    objLabInfo.Visible = True
    
    Call objImg.Refresh(False)
End Sub

Public Sub DrawErrorInfo(objImg As DicomImage, objImgInfo As clsBgImgInfo, _
    Optional ByVal blnIsClear As Boolean = False)
'���ƴ�����Ϣ
    Dim i As Long
    Dim objLabInfo As DicomLabel
    Dim objLabState As DicomLabel
    Dim lngLabIndex As Long
    Dim lngStateIndex As Long
    Dim lngBackIndex As Long
    Dim strErrorHint As String
    
    Set objLabInfo = Nothing
    Set objLabState = Nothing
    
    For i = 1 To objImg.Labels.Count
        If objImg.Labels(i).tag = IMG_LAB_ERRORINFO_TAG Then
            Set objLabInfo = objImg.Labels(i)
            lngLabIndex = i
        End If
        
        If objImg.Labels(i).tag = IMG_LAB_ERRORSTATE_TAG Then
            Set objLabState = objImg.Labels(i)
            lngStateIndex = i
        End If
    Next
    
    If blnIsClear Then
        If Not objLabInfo Is Nothing Then
            Call objImg.Labels.Remove(lngLabIndex)
        End If
        
        If Not objLabState Is Nothing Then
            Call objImg.Labels.Remove(lngLabIndex)
        End If
        
        If Not objLabInfo Is Nothing Or Not objLabState Is Nothing Then 'Or Not objLabBack Is Nothing
            Call objImg.Refresh(False)
        End If
        
        Exit Sub
    End If
    
    If objLabInfo Is Nothing Then
        Set objLabInfo = objImg.Labels.AddNew
        objLabInfo.tag = IMG_LAB_ERRORINFO_TAG
    End If
    
    If objImgInfo.Redo > 0 Then
        strErrorHint = "[��" & objImgInfo.Redo & "��" & IIf(objImgInfo.ImgCommand = icUpLoad, "�ϴ�", "����") & "]" & vbCrLf & objImgInfo.ErrorInfo
    Else
        strErrorHint = objImgInfo.ErrorInfo
    End If
    
    '����ı�������ͬ������Ҫˢ��
    If objLabInfo.Text = strErrorHint Then
        If (objImgInfo.LoadState = lsError) And Not (objLabState Is Nothing) Then Exit Sub
        If (objImgInfo.LoadState = lsRedo) Then Exit Sub
    End If
    
    If objImgInfo.LoadState = lsError Then
        If objLabState Is Nothing Then
            Set objLabState = objImg.Labels.AddNew
            objLabState.tag = IMG_LAB_ERRORSTATE_TAG
        End If
        
        objLabState.LabelType = doLabelText
        objLabState.FontSize = 10
'        objLabState.FontName = "����"
        objLabState.Font.Bold = True
        objLabState.AutoSize = False
       
        objLabState.Text = "!!" '"��"
        objLabState.Font.Bold = True
        objLabState.ForeColour = vbRed
        objLabState.BackColour = vbWhite
        objLabState.Shadow = doShadowAll
    
        objLabState.Transparent = False
        objLabState.ScaleWithCell = False
        objLabState.ImageTied = False
         
        objLabState.Left = 0
        
        
        objLabState.Top = 20 + (Len(objImgInfo.DrawHint) + 1) * 20 ' 21
       
       objLabState.Visible = True
    End If
    
    'Text*********************************************
    objLabInfo.LabelType = doLabelText
    objLabInfo.Margin = 0
    objLabInfo.FontSize = 10
    objLabInfo.AutoSize = True
    
    objLabInfo.Text = strErrorHint
    objLabInfo.ForeColour = vbRed
    objLabInfo.BackColour = vbYellow
    
    objLabInfo.Transparent = False
    objLabInfo.ScaleWithCell = False
     
    objLabInfo.Left = 40
    objLabInfo.Top = 1
    
    objLabInfo.Visible = True
    
    Call objImg.Refresh(False)
End Sub



Public Sub DrawBorder(objDcmImg As DicomImage, ByVal lngSelColorStyle As ColorConstants, _
    Optional ByVal blnIsSel As Boolean = False)
    Dim lngColor As OLE_COLOR
    
On Error GoTo errhandle
    lngColor = IMG_BACK_BORDER_COLOR
    If blnIsSel Then lngColor = lngSelColorStyle
 
    objDcmImg.BorderStyle = 0
    objDcmImg.BorderWidth = 1
    objDcmImg.BorderColour = lngColor
    
Exit Sub
errhandle:

End Sub


Public Sub DrawImgOrder(objDcmImg As DicomImage)
'����ͼ�����
    Dim objImgInfo As clsBgImgInfo
    Dim objLabOrder As DicomLabel
    
    Set objImgInfo = objDcmImg.tag
    If objImgInfo Is Nothing Then Exit Sub
    
    Set objLabOrder = objDcmImg.Labels.AddNew
    
    With objLabOrder
        .LabelType = doLabelText
        .tag = IMG_LAB_ORDER_TAG
    
        .FontSize = 9
        .AutoSize = False
        .Shadow = doShadowAll
 
        If objImgInfo.ImageOrder <= 0 Then
            .Text = "**"
        Else
            .Text = IIf(nvl(objImgInfo.SeriesNoTag) <= 1, "", objImgInfo.SeriesNoTag & "-") & IIf(objImgInfo.ImageOrder < 10, "0", "") & objImgInfo.ImageOrder
        End If
        
        .Font.Bold = True
        .ForeColour = &H40C0&
        .BackColour = vbWhite

        .Transparent = False
        .ScaleWithCell = False
     
        .Left = 0
        .Top = 21  '20 + 1 * 20
       
       
        .Visible = True
    End With
    
End Sub


Public Sub DrawCheckBox(objDcmImg As DicomImage, ByVal lngSelColorStyle As ColorConstants, _
    Optional ByVal blnIsSel As Boolean = False)
    Dim lSelect As DicomLabel
    Dim i As Long
    
On Error GoTo errhandle
    
    For i = 1 To objDcmImg.Labels.Count
        If objDcmImg.Labels(i).tag = IMG_LAB_CHECKBOX_TAG Then
            Set lSelect = objDcmImg.Labels(i)
            Exit For
        End If
    Next
 
    If lSelect Is Nothing Then
        Set lSelect = objDcmImg.Labels.AddNew
    Else
        If lSelect.Transparent = Not blnIsSel Then Exit Sub
    End If

    With lSelect
        .LabelType = doLabelRectangle            '����
        .Width = 18
        .Height = 18
        .Margin = 4
        .Left = 1
        .Top = 1
        .LineWidth = 2
        
'        .ForeColour = vbYellow
'        .BackColour = vbRed
        .ForeColour = CLng(&HC0C0C0)
        .BackColour = lngSelColorStyle 'CLng(&H8000000F)
        
        .Transparent = Not blnIsSel
        .ScaleWithCell = False
'        .ImageTied = False
    
        .tag = IMG_LAB_CHECKBOX_TAG
        
        .Visible = True
    End With
    
    Call objDcmImg.Refresh(False)
Exit Sub
errhandle:

End Sub


Public Sub DrawHints(objDcmImg As DicomImage)
    Dim i As Long
    Dim strHint As String
    Dim strChar As String
    
On Error GoTo errhandle
    
    strHint = objDcmImg.tag.DrawHint
    For i = 1 To Len(strHint)
        strChar = Mid(strHint, i, 1)
        Call DrawHint(objDcmImg, strChar)
    Next

Exit Sub
errhandle:


End Sub

Public Sub DrawHint(objDcmImg As DicomImage, ByVal strChar As String, _
    Optional ByVal blnIsClear As Boolean = False)
    Dim i As Long
    Dim objLabHint As DicomLabel
    Dim lngLabIndex As Long
    Dim lngHintIndex As Long
    
On Error GoTo errhandle
    
    Set objLabHint = Nothing
    lngHintIndex = 0
    
    For i = 1 To objDcmImg.Labels.Count
        If objDcmImg.Labels(i).tag = IMG_LAB_HINT_TAG Then
            lngHintIndex = lngHintIndex + 1
            If objDcmImg.Labels(i).Text = strChar Then
                Set objLabHint = objDcmImg.Labels(i)
                lngLabIndex = i
                
                Exit For
            End If
        End If
    Next
    
    If blnIsClear Then
        If objLabHint Is Nothing Then Exit Sub
        
            Call objDcmImg.Labels.Remove(lngLabIndex)
        
            '��Ҫ�ж��Ƿ����refresh
        Exit Sub
    End If
    
    If Not objLabHint Is Nothing Then Exit Sub  '�Ѿ��������˳�
    
    Set objLabHint = objDcmImg.Labels.AddNew
    
    objLabHint.LabelType = doLabelText
    objLabHint.tag = IMG_LAB_HINT_TAG
    
    objLabHint.FontSize = 10
    objLabHint.AutoSize = False
    objLabHint.Shadow = doShadowAll
 
    objLabHint.Text = strChar
    objLabHint.Font.Bold = True
    objLabHint.ForeColour = vbBlue
    objLabHint.BackColour = vbWhite

    objLabHint.Transparent = False
    objLabHint.ScaleWithCell = False
     
    objLabHint.Left = 0
    objLabHint.Top = 20 + (lngHintIndex + 1) * 20
       
       
    objLabHint.Visible = True
Exit Sub
errhandle:

End Sub


Public Sub DrawMarks(img As DicomImage, thisMarks As clsPicMarks, ByVal dblMarkZoom As Double)
'------------------------------------------------
'���ܣ���ʾ��ע��֧�����ֱ�ţ���ͷ��Բ�Σ����ֱ�ע
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim oneLabel As DicomLabel

    On Error GoTo err

    img.Labels.Clear
    'thisMarks(i).���Ͷ��� '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ), 6-˳���ţ�7-��ͷ��PACS�����ӣ�
    For i = 1 To thisMarks.Count
        With thisMarks(i)
            If thisMarks(i).���� = 0 Then       '�ı�
                img.Labels.Add GetNewLabel(doLabelText, .X1 * dblMarkZoom, .Y1 * dblMarkZoom, 0, 0)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.Font.Bold = True
                oneLabel.Text = .����
            ElseIf thisMarks(i).���� = 5 Then   '��Բ
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * dblMarkZoom, .Y1 * dblMarkZoom, (.X2 - .X1) * dblMarkZoom, (.Y2 - .Y1) * dblMarkZoom)
            ElseIf thisMarks(i).���� = 6 Then   '˳����
                'Բ�α���ɫ
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * dblMarkZoom - 8, .Y1 * dblMarkZoom - 8, 17, 17)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.XOR = False
                oneLabel.BackColour = IIf(.���ɫ = 0, vbYellow, .���ɫ)
                oneLabel.Transparent = False
'                oneLabel.tag = m_LabelTag_Back

                'Բ�ο�
'                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * dblMarkZoom - 8, .Y1 * dblMarkZoom - 8, 17, 17)
'                Set oneLabel = img.Labels(img.Labels.Count)
'                oneLabel.XOR = False
                oneLabel.ForeColour = vbBlack
'                oneLabel.Transparent = True
                oneLabel.tag = m_LabelTag_Circle
'                oneLabel.TagObject = img.Labels(img.Labels.Count - 1)


                'Բ�α������
                img.Labels.Add GetNewLabel(doLabelText, .X1 * dblMarkZoom - 8 + 2, .Y1 * dblMarkZoom - 8 + 2, 0, 0)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.ForeColour = vbBlack
                oneLabel.XOR = False
                oneLabel.Transparent = True
                oneLabel.tag = m_LabelTag_Number
                oneLabel.FontSize = 8
                oneLabel.FontName = "Arial Bold"
                oneLabel.AutoSize = True
                oneLabel.Text = .����
                If Val(.����) < 10 Then  '10���µ����֣���Ҫ΢��һ��λ�ã����ֲ��ܳ�����ԲȦ�����м�
                    oneLabel.Left = oneLabel.Left + 3
                End If

                oneLabel.TagObject = img.Labels(img.Labels.Count - 1)
                img.Labels(img.Labels.Count - 1).TagObject = oneLabel  'TagObject�γɱջ�  'img.Labels(img.Labels.Count - 2).TagObject = oneLabel

            ElseIf thisMarks(i).���� = 7 Then   '��ͷ
                img.Labels.Add GetNewLabel(doLabelArrow, .X1 * dblMarkZoom, .Y1 * dblMarkZoom, (.X2 - .X1) * dblMarkZoom, (.Y2 - .Y1) * dblMarkZoom)
            End If
        End With
    Next i
    
    Call img.Refresh(False)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub
