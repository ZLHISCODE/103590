Attribute VB_Name = "mdlPublic"
Option Explicit

Public gobjComLib As Object    'zl9ComLib.clsComLib
Public gcnOracle As ADODB.Connection
Public gcnOledb As ADODB.Connection
Public gstrPrivs As String
Public gstrSysName  As String
Public gstrDBUser As String
Public gstrSQL As String
Private mclsUnzip As Object
Public gobjPacsCore As Object   'PACS��Ƭ����

Public Const VIEW_ALLREPORT = "ȫԺӰ���ѯ"

Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function LoadImage Lib "SLInterCOM.dll" (ByVal hWnd As Long, ByVal pType As String, ByVal pStuNO As String, ByVal pParam1 As String, ByVal pParam2 As String, ByVal pParam3 As String) As Long
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type NETRESOURCE ' ������Դ
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Const RESOURCETYPE_ANY = &H0

Public Type tFtpInfo
    FtpDir As String
    FtpIP As String
    FtpPswd As String
    FTPUser As String
    DiviceId As String
    
    SubDir As String
    DestMainDir As String
End Type

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Function GetColNum(listTemp As Object, strHead As String) As Integer
    Dim i As Integer
    Select Case UCase(TypeName(listTemp))
        Case UCase("ReportControl")
            For i = 0 To listTemp.Columns.Count - 1
                If listTemp.Columns.Column(i).Caption = strHead Then GetColNum = listTemp.Columns.Column(i).ItemIndex: Exit Function
            Next
        Case UCase("ListView")
            For i = 1 To listTemp.ColumnHeaders.Count
                If listTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("MSHFlexGrid") '�������ʹ�������δ�õ�
        Case UCase("BillEdit")
        Case UCase("VSFlexGrid")
            For i = 0 To listTemp.Cols - 1
                If listTemp.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("BillEdit")
        Case UCase("DataGrid")
    End Select
End Function

Public Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function

Function GetFileContent(ByVal strFileName As String) As String
'��ȡ�����ļ�����
    Dim i As Integer, strContent As String, bty() As Byte
    
    If Dir(strFileName) = "" Then Exit Function
    
    i = FreeFile
    
    ReDim bty(FileLen(strFileName) - 1)
    
    Open strFileName For Binary As #i
    Get #i, , bty
    Close #i
    strContent = StrConv(bty, vbUnicode)
    
    GetFileContent = strContent
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '����������Դ
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "��������ʧ�ܣ��������������Ƿ���ȷ��"
    End If
    funcConnectShardDir = lngResult
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetAdviceID(ByVal lngReportID As Long) As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ҽ��ID from ����ҽ������ where ����ID =[1]"
    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡFTP��Ϣ", lngReportID)
    
    If rsData.RecordCount > 0 Then GetAdviceID = Val(Nvl(rsData!ҽ��ID))
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = IIf(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        GetUserInfo = True
    End If
End Function

Public Function View3DImage(ByVal lngҽ��ID As Long, frmParent As Object) As Long
    Dim blnCanViewImage As Boolean  '��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)ʱ���Ƿ���Թ�Ƭ
    Dim lngResut As Long
    Dim str3DViewType As String
    Dim intImageLocation As Long    'ͼ��λ�ã������������1���ɰ�PACS��2���ɰ�RIS+�°�PACS��3���°�RIS+PACS
    
    On Error GoTo DBError
    
    If getImageLocation(lngҽ��ID, intImageLocation, blnCanViewImage) = False Then Exit Function
    
    str3DViewType = gobjComLib.zlDatabase.GetPara("XW3D��Ƭ����", 100, 1288, "Study3D")
    If Trim(str3DViewType) = "" Then str3DViewType = "Study3D"
    
    lngResut = LoadImage(0, str3DViewType, CStr(lngҽ��ID), "", "", "")
    
    If lngResut = -121 Then
        MsgBox "���ò�������", vbInformation, gstrSysName
    ElseIf lngResut = -122 Or lngResut = -102 Then
        MsgBox lngResut & ":δ��ȷ��װPACS���ӿ��ļ�", vbInformation, gstrSysName
    ElseIf lngResut = -108 Then
        MsgBox lngResut & ":�������Ӵ���", vbInformation, gstrSysName
    ElseIf lngResut = -104 Then
        MsgBox lngResut & ":���ݿ����", vbInformation, gstrSysName
    ElseIf lngResut = -101 Then
        MsgBox lngResut & ":��������", vbInformation, gstrSysName
    End If
    
    View3DImage = lngResut
    
    Exit Function
DBError:
    lngResut = -1
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ViewStaticImage(ByVal lngҽ��ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
'���ܣ����ù�Ƭվ
    Dim intImageLocation As Long
    Dim blnCanViewImage As Boolean  '��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)ʱ���Ƿ���Թ�Ƭ
    
    On Error GoTo DBError
    
    If getImageLocation(lngҽ��ID, intImageLocation, blnCanViewImage, blnMoved) = False Then Exit Sub
    
    'ͼ�����������ݿ⣬�����������WEB���
    If intImageLocation = 1 Then
        Call XWWebViewerStaticOpen(lngҽ��ID)
    End If
    
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ViewPatientImage(ByVal lngҽ��ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ�����ҽ��ID�����ҵ����ߵ����в���ID����רҵ��PACS��Ƭ
'������lngҽ��ID--����ҽ��ID���
'       frmParent -- ������
'       blnMoved -- �Ƿ�ת��
'���أ���
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intImageLocation As Long
    Dim blnCanViewImage As Boolean  '��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)ʱ���Ƿ���Թ�Ƭ
    
    On Error GoTo DBError
    
    If getImageLocation(lngҽ��ID, intImageLocation, blnCanViewImage, blnMoved) = False Then Exit Sub
    
    'ͼ�����������ݿ⣬�����������WEB���
    If intImageLocation = 1 Then
        Call XWWebViewerPatientOpen(lngҽ��ID)
    End If
    
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ViewImage(ByVal lngҽ��ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
'���ܣ����ù�Ƭվ
    Dim strFtpHost As String
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    Dim intImageLocation As Long
    Dim lng����ID As Long
    Dim blnCanViewImage As Boolean  '��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)ʱ���Ƿ���Թ�Ƭ
    
    On Error GoTo DBError
    
    If getImageLocation(lngҽ��ID, intImageLocation, blnCanViewImage, blnMoved) = False Then Exit Sub
    
    'ͼ�����������ݿ⣬�����������WEB���
    If intImageLocation = 1 Or intImageLocation = 2 Then
        Call XWWebViewerOpen(lngҽ��ID)
        
        If intImageLocation = 2 Then
            Call XWDownLoadImage(lngҽ��ID)
        End If
        
        Exit Sub
    End If
    
    
    '���ж��Ƿ����ͼ��û��ͼ������ʾ���˳�
    strSql = "Select A.���UID,Count(B.����UID) as �������� From Ӱ�����¼ A,Ӱ�������� B Where A.���UID=B.���UID And A.ҽ��ID=[1] Group by A.���UID"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��Ƭ����", lngҽ��ID)
    If rsTmp.EOF Then
        MsgBox "û�п����ڹ�Ƭ�ı���ͼ��", vbInformation, gstrSysName
        Exit Sub
    End If

    strFtpHost = ""
    
    '������Ҫ�򿪵�����ͼ����Ϣ
    strSql = "Select /*+RULE*/ D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'\')" & _
        "||C.���UID||'\' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
        
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����Ŀ¼��Ϣ", lngҽ��ID)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        ClearCacheFolder App.Path & "\TmpImage\"
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strFtpHost = rsTmp("Host1")
            strSDPath = Nvl(rsTmp("����Ŀ¼1"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����1"))
        ElseIf Nvl(rsTmp("�豸��2")) <> "" Then
            strFtpHost = rsTmp("Host2")
            strSDPath = Nvl(rsTmp("����Ŀ¼2"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        On Error Resume Next
        If strSDPath <> "" Then
            Call funcConnectShardDir("\\" & strFtpHost & "\" & strSDPath, strSDUser, strSDPwd)
        End If
        
        If gobjPacsCore Is Nothing Then
            Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
        End If
        gobjPacsCore.CallOpenViewer "", lngҽ��ID, frmParent, gcnOracle, blnMoved, False
        
    End If

    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function CheckEPRReport(ByVal lngҽ��ID As Long, Optional lng����ID As Long, Optional blnBySign As Boolean, Optional ByVal intִ��״̬ As Integer = -999) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lngҽ��ID=�ɼ��е�ҽ��ID
'      lng����ID=���Դ��룬��Ҫ���ڷ��ر��没��ID
'      intִ��״̬=���ڼ������ʱ�������ۺϵ�ִ��״̬
'������blnBySign=�����Ƿ����ͨ��ǩ�������ж�(����ҽ������վ)
'���أ�0-���滹û����д
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrH
    
    '��鱨���Ƿ�����д
    If lng����ID = 0 Then
        strSql = "Select ����ID From ����ҽ������ Where ҽ��ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID)
        If Not rsTmp.EOF Then lng����ID = rsTmp!����id
    End If
    If lng����ID = 0 Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
        '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
        strSql = _
            " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
            " Union ALL" & _
            " Select ����,ҽ��ID,ִ�й���,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
            " From (" & _
                " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                    " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                " Order by A.���" & _
            " ) Where Rownum=1" & _
            " Order by ����,����ʱ�� Desc"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID, intִ��״̬)
        If Nvl(rsTmp!ִ�й���, 0) >= 5 Or Nvl(rsTmp!ִ��״̬, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        'ͨ��ǩ���汾�жϱ�����ɵķ�ʽ
        strSql = "Select B.�ļ�ID,Max(B.��ʼ��) as ǩ���汾 From ���Ӳ������� B Where B.�ļ�ID=[1] And B.��������=8 Group by B.�ļ�ID"
        strSql = "Select B.���ʱ��,B.���汾,C.ǩ���汾 From ���Ӳ�����¼ B,(" & strSql & ") C Where B.ID=[1] And B.ID=C.�ļ�ID(+)"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng����ID)
            
        '(ǩ������ֱ���޸ģ������޶������ǩ�������汾Ӧ��ǩ���汾һ��)
        If IsNull(rsTmp!���ʱ��) Or Nvl(rsTmp!���汾, 0) <> Nvl(rsTmp!ǩ���汾, 0) Then
            '���ҽ�������Ѿ�ִ��,��ʹû��ǩ���򲻷�Ҳ��ͬ���
            strSql = _
                " Select 2 as ����,ҽ��ID,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
                " Union ALL" & _
                " Select ����,ҽ��ID,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
                " From (" & _
                    " Select 1 as ����,B.ҽ��ID,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                    " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                        " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                    " Order by A.���" & _
                " ) Where Rownum=1" & _
                " Order by ����,����ʱ�� Desc"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID, intִ��״̬)
            If Nvl(rsTmp!ִ��״̬, 0) = 1 Then
                CheckEPRReport = 1
            Else
                CheckEPRReport = 2
            End If
        Else
            CheckEPRReport = 1
        End If
    End If
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function XWDownLoadImage(lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ����ƽ̨����ͼ��
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strStudyUID As String
    
    On Error GoTo err
    strSql = "select ���UID from Ӱ�����¼ where ҽ��ID = [1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ���UID", lngOrderID)
    If rsTemp.EOF = True Then Exit Function
    
    strStudyUID = Nvl(rsTemp!���UID, "")
    
    '���������洢���̡�P_OEM_DOWNLOADIMG_RIS��������ƽ̨����ͼ��
    strSql = "P_OEM_DOWNLOADIMG_RIS@XWPacs('" & strStudyUID & "')"
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSql, gstrSysName)
    
    Call MsgBox("�û���Ӱ��ͼƬ�Ѿ��ϴ��ƶˣ�����ƶ����أ�" & vbCrLf & vbCrLf & "�˹�����Ҫһ��ʱ�䣬" & vbCrLf & vbCrLf & "��û����ͼƬ˵���������أ����Եȡ�", vbOKOnly, "��ʾ��Ϣ")
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then Resume
    XWDownLoadImage = 1
End Function

Private Function XWWebViewerOpen(lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    
    On Error GoTo err
    
    strPath = gobjComLib.zlDatabase.GetPara("XWWEB��Ƭ��ַ", 100, 1288, "")
    
    If strPath <> "" Then
        strPath = Replace(strPath, "[@STU_NO]", lngOrderID)

        '����64λ�Ĳ���ϵͳ��XW WEB��Ƭ��֧��64λIE������Ҫʹ��32λ��IE
        If Dir("C:\Program Files (x86)\Internet Explorer", vbDirectory) = "" Then
            strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        Else
            strURL = "C:\Program Files (x86)\Internet Explorer\iexplore.exe " & strPath
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        MsgBox "XWWEB��Ƭ��ַΪ�գ��������ú�WEB��������", vbOKOnly, "��ʾ��Ϣ"
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function XWWebViewerStaticOpen(lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    
    On Error GoTo err
    
    strPath = gobjComLib.zlDatabase.GetPara("XW�ؼ�ͼ���ַ", 100, 1288, "")
    
    If strPath <> "" Then
        strPath = Replace(strPath, "[@STU_NO]", lngOrderID)
        
        '����64λ�Ĳ���ϵͳ��XW WEB��Ƭ��֧��64λIE������Ҫʹ��32λ��IE
        If Dir("C:\Program Files (x86)\Internet Explorer", vbDirectory) = "" Then
            strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        Else
            strURL = "C:\Program Files (x86)\Internet Explorer\iexplore.exe " & strPath
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerStaticOpen = 0
    Else
        MsgBox "XW�ؼ�ͼ���ַַΪ�գ��������úùؼ�ͼ���ַ��", vbOKOnly, "��ʾ��Ϣ"
        XWWebViewerStaticOpen = 1
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function XWWebViewerPatientOpen(lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������WEB Viewer�����ݲ���ID����ʾ����б���Ƭ
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    Dim strPatientIDs As String     '����רҵ��PACS�Ĳ���ID������ʽ�ǡ�'���1','���2'��
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strPath = gobjComLib.zlDatabase.GetPara("XWWeb����б��Ƭ��ַ", 100, 1288, "")
    
    If strPath <> "" Then
    
        '����ҽ��ID����ȡ����ID��
        strSql = "select ����ID from ������Ϣ  where ���֤��=(select ���֤�� from ������Ϣ a,����ҽ����¼ b where a.����id =b.����id and b.id=[1]) "
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����ID��", lngOrderID)
        If rsTemp.EOF = True Then
            strSql = "select ����ID from ����ҽ����¼  where id=[1] "
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����ID��", lngOrderID)
            If rsTemp.EOF = True Then
                MsgBox "����ҽ��ID " & lngOrderID & " ��ȡ��������ID��", vbOKOnly, "zlPublicPACS��Ƭ��ʾ"
                Exit Function
            End If
        End If
        
        While rsTemp.EOF = False
            strPatientIDs = strPatientIDs & ",'" & rsTemp!����ID & "'"
            rsTemp.MoveNext
        Wend
        strPatientIDs = Mid(strPatientIDs, 2)
        
        strPath = Replace(strPath, "[@PAT_NOs]", strPatientIDs)
        
        '����64λ�Ĳ���ϵͳ��XW WEB��Ƭ��֧��64λIE������Ҫʹ��32λ��IE
        If Dir("C:\Program Files (x86)\Internet Explorer", vbDirectory) = "" Then
            strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        Else
            strURL = "C:\Program Files (x86)\Internet Explorer\iexplore.exe " & strPath
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerPatientOpen = 0
    Else
        MsgBox "XWWeb����б��Ƭ��ַΪ�գ��������ù�Ƭ��ַ��", vbOKOnly, "��ʾ��Ϣ"
        XWWebViewerPatientOpen = 1
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub BlobToFile(fld As ADODB.Field, Filename As String, Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesleft As Long, bytes As Long
    Dim tmp() As Byte
    
    If (fld.Attributes And adFldLong) = 0 Then
        err.Raise 1001, , "field doesn't support the GetChunk method."
    End If
    
    If Dir$(Filename) <> "" Then Kill Filename
    
    fnum = FreeFile
    Open Filename For Binary As fnum
    bytesleft = fld.ActualSize
    Do While bytesleft
        bytes = bytesleft
        If bytes > ChunkSize Then bytes = ChunkSize
        tmp = fld.GetChunk(bytes)
        Put #fnum, , tmp
        bytesleft = bytesleft - bytes
    Loop
    
    Close #fnum
End Sub

Public Function InitOledbConn(Optional ByVal blnUseAlone As Boolean = False) As Boolean
    Dim objRegister As Object
    Dim strError As String

On Error GoTo err

    If blnUseAlone Then
        Set objRegister = GetObject("", "zlRegisterAlone.clsRegister")
    Else
        Set objRegister = GetObject("", "zlRegister.clsRegister")
    End If

    Set gcnOledb = objRegister.ReGetConnection(1, strError)

    InitOledbConn = True
    Exit Function
err:
    InitOledbConn = False

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetRecordset(ByVal strSql As String) As ADODB.Recordset
On Error GoTo ErrHand
    If gcnOledb Is Nothing Then
        Call InitOledbConn
    End If

    Set GetRecordset = New ADODB.Recordset

    If gcnOledb Is Nothing Then Exit Function

    If GetRecordset.State = adStateOpen Then GetRecordset.Close
    '��
    GetRecordset.Open strSql, gcnOledb, adOpenKeyset, adLockOptimistic
     
'    Set GetRecordset = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSql, "�ж��Ƿ���Ӱ��ͼƬ", Null, 1)
 
    Exit Function
ErrHand:
    If err <> 0 Then
        MsgBox "��������" & err.Description, vbInformation, "ϵͳ��Ϣ"
    End If
End Function

Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
End Sub

Public Function HasImage(lngOrderID As Long) As Boolean
''--------------------------------------------
''���ܣ� �жϸü���Ƿ���ͼ��
'           lngAdviceID -- ҽ��ID
''���أ�True-��ͼ��False-��ͼ��
''--------------------------------------------
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intImageLocation As Integer

    On Error GoTo err
    
    '�жϸü���Ƿ���ͼ�������������1���ɰ�PACS��2���ɰ�RIS+�°�PACS��3���°�RIS+PACS
    
    '�Ȳ�ѯͼ���Ƿ��ھɰ�PACS
    strSql = "Select ���UID,ͼ��λ�� From Ӱ�����¼ Where ҽ��ID =[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "�ж��Ƿ���Ӱ��ͼƬ", lngOrderID)
    
    intImageLocation = 0
    If rsTemp.RecordCount > 0 Then
        'ͼ���������1��2
        If Nvl(rsTemp!ͼ��λ��, 0) = 0 Then
            'ͼ���ھɰ�PACS��
            '����� ���UID �ļ�¼˵�����ݿ�����ͼ���򷵻�True����֮����false
            HasImage = IIf(Nvl(rsTemp!���UID, 0) <> 0, True, False)
        Else
            'ͼ���ھɰ�RIS+�°�PACS��
            intImageLocation = 1
        End If
    Else
        '����255��������ʹ����Ӱ����Ϣϵͳרҵ�棬ͼ�����°�RIS+PACS��
        If Val(gobjComLib.zlDatabase.GetPara(255, 100)) = 1 Then
            intImageLocation = 1
        End If
    End If
    
    If intImageLocation = 1 Then
        'ͼ�����°�PACS��,���� ִ�й���>=3 �ж��Ƿ���ͼ��
        strSql = "SELECT ҽ��ID from ����ҽ������  where ִ�й���>=3 and ҽ��ID =[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "�ж��Ƿ���Ӱ��ͼƬ", lngOrderID)
        
        If rsTemp.EOF Then
            HasImage = False
        Else
            HasImage = True
        End If
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isUseXWInterface(strSubName As String) As Boolean
''--------------------------------------------
''���ܣ� �ж��Ƿ�ʹ������RIS
'           strSubName -- ���õĳ�������
''���أ�True-ʹ�ã�False-��ʹ��
''--------------------------------------------
    Dim strUseXWInterface As String
    
    On Error GoTo err
    
    strUseXWInterface = gobjComLib.zlDatabase.GetPara(255, 100)
    
    BUGEX strSubName & ": strUseXWInterface = " & strUseXWInterface
    
    '��ȡ�Ƿ�����Ӱ����Ϣϵͳ�ӿ�
    isUseXWInterface = Val(strUseXWInterface) = 1
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function getImageLocation(ByVal lngҽ��ID As Long, ByRef intImageLocation As Long, ByRef blnCanViewImage As Boolean, _
    Optional ByVal blnMoved As Boolean = False) As Boolean
''--------------------------------------------
''���ܣ� �ж�ͼ��λ�ã��Ƿ����δ��˱���鿴ͼ��
''������    lngҽ��ID -- ҽ��ID
'           intImageLocation -- ͼ��λ�ã������������1���ɰ�PACS intImageLocation=0��
'                   2���ɰ�RIS+�°�PACS intImageLocation=1��2(�ϴ����ƴ洢)��3���°�RIS+PACS intImageLocation=1
'           blnCanViewImage -- ��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)ʱ���Ƿ���Թ�Ƭ
''���أ�True-�ɹ���False-ʧ��
''--------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lngִ�п���ID As Long
    Dim blnIsGreen As Boolean
    Dim blnIsUrgent As Boolean
    
    On Error GoTo err
    
    lngִ�п���ID = 0

    '��ѯͼ��λ��,�Լ�ִ�п���ID
    strSql = "Select a.ͼ��λ��, a.ִ�п���id, a.��ɫͨ��, b.������־ From Ӱ�����¼ a, ����ҽ����¼ b Where a.ҽ��id = b.Id And a.ҽ��id =[1]"
    
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
    End If
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ѯͼ�����ڵ�λ��", lngҽ��ID)
    
    If rsTmp.RecordCount <> 0 Then
        intImageLocation = Nvl(rsTmp!ͼ��λ��, 0)
        lngִ�п���ID = Val(Nvl(rsTmp!ִ�п���ID, 0))
        blnIsGreen = IIf(Val(Nvl(rsTmp!��ɫͨ��, 0)) = 1, True, False)
        blnIsUrgent = IIf(Val(Nvl(rsTmp!������־, 0)) = 1, True, False)
    Else
        intImageLocation = 1
    End If
    
    If lngִ�п���ID > 0 Then
        'ͼ�����λ��1��2
        strSql = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ����", lngִ�п���ID, "��ͼ��ҽ��վ���ɹ�Ƭ")
        If rsTmp.RecordCount > 0 Then blnCanViewImage = Val(Nvl(rsTmp!����ֵ, 0)) = 1
    Else
        'ͼ�����λ��3������ҽ��ID�������
        blnCanViewImage = isUseXWInterface("getImageLocation")
    End If
    
    '��ȡ����״̬
    strSql = "Select ִ�й��� from ����ҽ������ where ҽ��id= " & lngҽ��ID
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "getImageLocation", lngҽ��ID)
    
    If rsTmp.RecordCount > 0 Then
        If blnCanViewImage Then
            '�������δ��ɣ����ҹ�ѡ�˲�������ͼ��ҽ��վ���ɹ�Ƭ���������С����ǰ��Ƭ��Ȩ��ʱ�ſɽ��й�Ƭ
            '�������ɫͨ�����˲��������ǰ��ƬȨ��
            If Nvl(rsTmp!ִ�й���, 0) < 5 Then
                If InStr(gstrPrivs, "���ǰ��Ƭ") <= 0 And Not (blnIsGreen And blnIsUrgent) Then
                    MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û�����ǰ��ƬȨ��ʱ���ܲ鿴ͼ��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            'û�й�ѡ��������ͼ��ҽ��վ���ɹ�Ƭ��ʱ��������ɺ�ſɽ��й�Ƭ
            If Nvl(rsTmp!ִ�й���, 0) < 5 Then
                MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)�����ڲ��ܲ鿴ͼ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    getImageLocation = True
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
