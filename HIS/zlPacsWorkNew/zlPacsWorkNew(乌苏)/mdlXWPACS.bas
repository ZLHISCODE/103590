Attribute VB_Name = "mdlXWPACS"
Option Explicit

Public Const SW_HIDE = 0
Public Const SW_SHOWMAXIMIZED = 3
Public Const WM_USER = &H400



Public gcnXWDBServer As New ADODB.Connection         '�������ݿ�����
Public gfrmPacsMain As frmPacsMain                  '�������ձ���ͼ��Ϣ�Ĵ���ָ��
Public glngXWDeptID As Long                         '��ǰ����ID
Public gblnXWMoved As Boolean                       '�Ƿ�ת��
Public gblnXWLog As Boolean                         '�Ƿ��¼ͨѶ��־
Public gstrOracleOwner As String                    'oracle����ӵ����
Public gblnUseXinWangView As Boolean                '�ж��Ƿ�������������Ƭ
Public gstrImageShareDir As String                  '�ϰ��Ӱ����洢Ŀ¼
Public glngStudySchemeNo As Long                    '��鷽����
Public glngSeriesSchemeNo As Long                   '���з�����


'���������ṩ�� InterCOM.dll������ADViewer�鿴ͼ��

'��������
Public Declare Function OEMViewStart Lib "InterCOM.dll" (ByVal cpReserved1 As String, ByVal cpReserved2 As String, ByVal cpReserved3 As String) As Long
Public Declare Function OEMViewExit Lib "InterCOM.dll" (ByVal cpReserved1 As String, ByVal cpReserved2 As String, ByVal cpReserved3 As String) As Long
Public Declare Function OEMViewOpen Lib "InterCOM.dll" (ByVal lPlanID As Long, ByVal cpFilter As String, ByVal lFunc As Long, ByVal cpReserved As String) As Long
Public Declare Function OEMViewClose Lib "InterCOM.dll" (ByVal cpReserved As String) As Long
Public Declare Function LoadImage Lib "SLInterCOM.dll" (ByVal hWnd As Long, ByVal pType As String, ByVal pStuNO As String, ByVal pParam1 As String, ByVal pParam2 As String, ByVal pParam3 As String) As Long

'��ƬԤ��
Private Declare Function SmitPrintFilm Lib "SLPrnDicomFilm.dll" (ByVal lNoType As Long, ByVal cpNo As String, ByVal lFlag As Long, _
    ByVal cpUserName As String, ByVal cpStationNo As String, ByVal cpDevGName As String, ByVal cpDevCName As String, _
    ByVal cpReserve As String, ByVal cpPatInfoFile As String) As Long
    
'��Ƭ��ӡ
Private Declare Function PreviewFilmEx Lib "SLPrnDicomFilm.dll" (ByVal lNoType As Long, ByVal cpNo As String, _
    ByVal cpUserName As String, ByVal cpStationNo As String, ByVal cpReserve As String) As Long
 
 
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

 
 
'�����ṩ�ĺ���
'����ADViewer�� long OEMViewStart ( LPCTSTR cpReserved1, LPCTSTR cpReserved2, LPCTSTR cpReserved3 );
'�˳�ADViewer�� long OEMViewExit ( LPCTSTR cpReserved1, LPCTSTR cpReserved2, LPCTSTR cpReserved3 );
'��ָ��ͼ�� long OEMViewOpen ( long lPlanID, LPCTSTR cpFilter, long lFunc, LPCTSTR cpReserved );
'�ر�ͼ�� long OEMViewClose ( LPCTSTR cpReserved );


'���ձ���ͼ����Ϣ��API
Public Const WM_XWReportImage As Long = 5120
'��ϢHook����
Public plngXWPreWndProc As Long       'ԭ������Ϣ�������


'�ж��Ƿ����ù�Ƭ
Function IsUseXwViewer() As Boolean
On Error GoTo ErrHandle
    Dim lngPhkResult As Long
    Dim lngKey As Long
    Dim blnResult As Boolean
    Dim strValue As String
    Dim lngLen As Long
    
    blnResult = IIf(RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Silver\Silver Pacs\General", 0, 1, lngPhkResult) = 0&, True, False)
    
    '�����װ��xw viewer����ע����еİ�װ��Ϣ
    If blnResult = False Then Exit Function
    
    strValue = Space(255)
    lngLen = Len(strValue)
    
    Call RegQueryValueEx(lngPhkResult, "CompanyName", 0, 0, strValue, lngLen)
    
    Call RegCloseKey(lngKey)
    
    If InStr(strValue, "����ҽѧӰ��ϵͳ") <= 0 Then blnResult = False
    
    IsUseXwViewer = blnResult
Exit Function
ErrHandle:
    IsUseXwViewer = False
End Function

'-----------------------------------------------------------------------------------------------------
'ADViewer��������
'-----------------------------------------------------------------------------------------------------
Function XWFilmPreview(ByVal strOrderNo As Long) As Long
'Ԥ���������н�Ƭ

'PreviewFilmEx����˵��
'lNoType:Ĭ��2015����ʾ���ݼ��ҽ��Ԥ��
'cpNo:��������
'cpUserName:��Ƭ��ӡ���û���,���Թ̶�����Ϊris
'cpStationNo:�ɲ�����
'cpReserve:����

    XWFilmPreview = PreviewFilmEx(2015, strOrderNo, "ris", "", "")
End Function


Function XWFilmPreviewWithFile(ByVal strFilmPath As String) As Long

'PreviewFilmEx����˵��
'lNoType:Ĭ��2015
'cpNo:��������
'cpUserName:��Ƭ��ӡ���û���,���Թ̶�����Ϊris
'cpStationNo:�ɲ�����
'cpReserve:����

    XWFilmPreviewWithFile = Shell("c:\ris\FilmPreview.exe " & strFilmPath, vbHide)
End Function

Function XWFilmPreviewEx(ByVal lngFilmId As Long) As Long
'��ƬԤ��

'PreviewFilmEx����˵��
'lNoType:Ĭ��2015�� 2000��ʾ���ݽ�ƬID����Ԥ��
'cpNo:��������
'cpUserName:��Ƭ��ӡ���û���,���Թ̶�����Ϊris
'cpStationNo:�ɲ�����
'cpReserve:����
On Error Resume Next
    
    XWFilmPreviewEx = PreviewFilmEx(2000, lngFilmId, "ris", "", "+ShowMessage")
    If XWFilmPreviewEx <> 0 Then
        XWFilmPreviewEx = PreviewFilmEx(2000, lngFilmId, "ris", "", "+ShowMessage")
    End If

    'XWFilmPreviewEx = RunDll32("c:\RIS\SLPrnDicomFilm.dll", "PreviewFilmEx", 2000, lngFilmId, "ris", "tjh", "+ShowMessage")
End Function

Function XWFilmPrint(ByVal strOrderNo As String, ByVal lngPrintType As Long, ByVal strPrintName As String) As Long
'��Ƭ��ӡ
'lngOrderID:ҽ��ID
'lngPrintType:������ͣ�0-�ڰ�,1-��ɫ
'strPrintName:��ӡ������

    If lngPrintType = 0 Then
        XWFilmPrint = SmitPrintFilm(2015, strOrderNo, 0, "ris", "", strPrintName, "", "+ShowMessage", App.Path & "\FilmPrintResult.ini")
    Else
        XWFilmPrint = SmitPrintFilm(2015, strOrderNo, 0, "ris", "", "", "strPrintName", "+ShowMessage", App.Path & "\FilmPrintResult.ini")
    End If
End Function

Function XWFilmDelete(ByVal lngFilmId As Long) As Boolean
'ɾ����Ƭ
'lngFilmId:��ƬID
    Dim strSQL As String
    
    strSQL = "ris.p_oem_del_film(" & lngFilmId & ")"
    Call gcnXWDBServer.Execute(strSQL)
    
    XWFilmDelete = True
End Function



Function XWADViewerStart() As Long
'--------------------------------------------
'���ܣ� ����ADViewer
'       �ú���ͨ��ֻ��Ҫ����һ�Ρ���Ȼ��ͼ��ʱ���ADViewer ���Զ�������
'       ����Ϊ�˼ӿ�ִ���ٶȣ�������������������ʱִ�д˺�������ͬʱ����ADViewer
'��������
'���أ�
'--------------------------------------------
    'OEMViewStart ���������� cpReserved1��cpReserved2��cpReserved3����Ϊ�������̶�ΪNULL
    
    On Error GoTo err
    
    XWADViewerStart = OEMViewStart("", "", "")
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Function XWADViewerExit() As Long
'--------------------------------------------
'���ܣ� �˳�ADViewer
'
'��������
'���أ�
'--------------------------------------------
    'XWViewerExit ���������� cpReserved1��cpReserved2��cpReserved3����Ϊ�������̶�ΪNULL
    
    On Error GoTo err
    
    XWADViewerExit = OEMViewExit("", "", "")
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Function XWADViewerOpen(ByVal strFilter As String, ByVal lngPlanId As Long) As Long
'--------------------------------------------
'���ܣ� ��ָ��ͼ�������ɲ���ָ�����ұ����������ļ��е����������
'       ͼ���ʱ��ADViewer ��ǰģʽ�йأ�����ǵ���¼ģʽ����������Զ��ر�ԭ����ͼ������ǶԱ�
'       ģʽ�������ӵ�ADViewer�С�
'������
'       lngOrderID -- ҽ��ID
'���أ�
'--------------------------------------------
    Dim strRev As String
    Dim lngFunction As Long
    Dim strXwPrivs As String
    
    'XWViewerOpen ����˵����
    'lPlanID��  ����ID����ID ������INI �ļ���һ�£��ڼ����������£�ͨ����ֵΪ1������Ѹ�ID ��Ϊһ���������ʱ��ȡ������롣
    'cpFilter�� �ô���Ҫ��ͼ�������ֵ��������š�����ŵȣ����Դ�����ֵ��
    '           ��ֵ֮ͬ���÷ָ���[;]�������ò��������弰˳����INI �ļ������ã�������lPlanID��Ӧ��
    'lFunc��    ����Ȩ�ޡ�ÿһλ����һ��ܣ�������ж���Ȩ�ޣ���λ���򡱼��ɣ����幦������:
    '           0x00000002�� �ؽ�ͼ�񱣴棬���磺��Ӱ��ͼ��ƴ��ͼ���
    '           0 x00000200: ��Ƭ��ӡ
    '           0 x00040000: ͼ�񵼳�?���Ϊ������ʽ
    '           0 x00080000: GSPS ����
    'cpReserved��   ��������ΪNULL
    
    On Error GoTo err
    
    '��¼�ӿ���־
    If gblnXWLog = True Then
        Call WriteCommLog("XWADViewerOpen", "XW�ӿ�", "��ADViewer������ʾͼ��ҽ��ID= " & strFilter)
    End If
    
    '����RIS�е�Ȩ�ޣ���֯Ȩ�޴�
    lngFunction = 0
    strXwPrivs = GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    If InStr(strXwPrivs, "PACS�����ؽ�ͼ��") <> 0 Then
        lngFunction = lngFunction Or &H2
    End If
    
    If InStr(strXwPrivs, "PACS��Ƭ��ӡ") <> 0 Then
        lngFunction = lngFunction Or &H200
    End If
    
    If InStr(strXwPrivs, "PACSͼ�񵼳�") <> 0 Then
        lngFunction = lngFunction Or &H40000
    End If
    
    If InStr(strXwPrivs, "PACS GSPS����") <> 0 Then
        lngFunction = lngFunction Or &H80000
    End If
        
    XWADViewerOpen = OEMViewOpen(lngPlanId, strFilter, lngFunction, "")
    
    If XWADViewerOpen <> 0 Then
        MsgBox "ADViewer�򿪴��󣬷��ص���Ϣ�ǣ�" & XWADViewerOpen
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Function XWADViewerClose() As Long
'--------------------------------------------
'���ܣ� �ر�ͼ�񣬲��˳�ADViewer
'
'��������
'���أ�
'--------------------------------------------
    'XWViewerClose �Ĳ��� cpReserved1Ϊ�������̶�ΪNULL
    On Error GoTo err
    
    XWADViewerClose = OEMViewClose("")
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As ADODB.Connection
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOra As New ADODB.Connection
    
    On Error Resume Next
    err = 0
    
    DoEvents
    
    With cnOra
        If .State = adStateOpen Then .Close
        
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            'OraDataOpen = Nothing
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    Set OraDataOpen = cnOra
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function


Sub XWTestDBConnection(ByVal strServerName As String, ByVal strUser As String, ByVal strPwd As String)
'���ܣ� ��������SQLServer���ݿ�����
'������
'���أ��ɹ����ؿ��ַ�
'--------------------------------------------
    Dim cnTest As New ADODB.Connection

    If strServerName = "" Then
        MsgBox "δ�ҵ����ݿ������������Ϣ�������á�"
        Exit Sub
    End If
    
    On Error Resume Next
    err = 0
    
    If cnTest.State = adStateOpen Then cnTest.Close
    
    Set cnTest = OraDataOpen(strServerName, strUser, strPwd)
    
    If err <> 0 Or cnTest Is Nothing Then
        '���ݿ����Ӵ���
        MsgBox "���ݿ�����ʧ�ܡ�" & vbCrLf & vbCrLf & "��������ǣ�" & err.Number & "�����������ǣ� " & err.Description
        Exit Sub
    End If
    
    MsgBox "���ݿ����ӳɹ���"
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


'----------------------------------------------------------------------------------------------
'����SQLSERVER���ݿ����Ӻ͹ر�

'----------------------------------------------------------------------------------------------

Public Function XWDBServerOpen() As Boolean
'--------------------------------------------
'���ܣ� ��������SQLServer���ݿ�
'��������
'���أ�0-�ɹ�
'--------------------------------------------
    Dim strSqlUser As String
    Dim strSqlPWD As String
    Dim strDataSource As String

    XWDBServerOpen = False
    
    If InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "����") <= 0 Then Exit Function
    
    '������ORACLE ģ������л�ȡ���������ݿ������IP��ַ���û���������
    strDataSource = zlDatabase.GetPara("XW���ݿ������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    strSqlUser = zlDatabase.GetPara("XW���ݿ�������û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    strSqlPWD = zlDatabase.GetPara("XW���ݿ����������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    gstrImageShareDir = zlDatabase.GetPara("XW��ʷͼ����Ŀ¼", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    glngStudySchemeNo = Val(zlDatabase.GetPara("XW��鷽����", glngSys, G_LNG_XWPACSVIEW_MODULE, "1"))
    glngSeriesSchemeNo = Val(zlDatabase.GetPara("XW���з�����", glngSys, G_LNG_XWPACSVIEW_MODULE, "2"))
    
    If strDataSource = "" Then
        MsgBox "δ�ҵ�SQLSERVER���ݿ�����������ڡ�Ӱ��RIS����վ����PACS���������á�"
        XWDBServerOpen = 1
        Exit Function
    End If

    On Error Resume Next
    err = 0
    If gcnXWDBServer.State = adStateOpen Then gcnXWDBServer.Close
    
    Set gcnXWDBServer = OraDataOpen(strDataSource, strSqlUser, strSqlPWD)
    
    If err <> 0 Or gcnXWDBServer Is Nothing Then
        '���ݿ����Ӵ���
        MsgBox "DBServer���ݿ����Ӵ��󣬿��ܻᵼ�²���ͼ���޷��鿴��" & vbCrLf & vbCrLf & "��������ǣ�" & err.Number & "�����������ǣ� " & err.Description
    End If
    
    XWDBServerOpen = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Function XWDBServerClose() As Long
'--------------------------------------------
'���ܣ� �ر�����SQLServer���ݿ�����
'��������
'���أ�0-�ɹ�
'--------------------------------------------
    On Error GoTo err
    
    If gcnXWDBServer Is Nothing Then Exit Function
    
    If gcnXWDBServer.State = adStateOpen Then gcnXWDBServer.Close
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function XWShow3DImage(ByVal lngAdviceID As Long, objParent As Object)
'���ܣ�3D��Ƭ
'������
'   lngStudyNo---����

    On Error GoTo err
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWShow3DImage", "XW�ӿ�", "����3D��Ƭ")
    End If
    
    Call XWWeb3DViewerOpen(lngAdviceID, objParent)
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWWeb3DViewerOpen(ByVal lngAdviceID As Long, objParent As Object) As Long
'���ܣ� ��������3D��Ƭ
'0   �ɹ�
'-121    ���ò�������
'-122    δ��ȷ��װPACS���ӿ��ļ�
'-102    δ��ȷ��װPACS���ӿ��ļ�
'-103    �������Ӵ���
'-104    ���ݿ����
'-101    ��������
    Dim str3DViewType As String
    
    On Error GoTo err
    
    If lngAdviceID <= 0 Then
        '��¼�ӿ���־
        If gblnXWLog = True Then
            Call WriteCommLog("XWShowImage", "XW�ӿ�", "��������3D��Ƭ,ҽ��IDΪ��")
        End If
        
        XWWeb3DViewerOpen = -101
        Exit Function
    Else
        '��¼�ӿ���־
        If gblnXWLog = True Then
            Call WriteCommLog("XWShowImage", "XW�ӿ�", "��������3D��Ƭ,ҽ��IDΪ:" & lngAdviceID)
        End If
        
        str3DViewType = zlDatabase.GetPara("XW3D��Ƭ����", glngSys, G_LNG_XWPACSVIEW_MODULE, "Study3D")
        If Trim(str3DViewType) = "" Then str3DViewType = "Study3D"
        
        XWWeb3DViewerOpen = LoadImage(objParent.hWnd, str3DViewType, CStr(lngAdviceID), "", "", "")
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    XWWeb3DViewerOpen = -101
End Function


'-------------------------------------------------------------------------------------------------------
'ADViewer�鿴ͼ��Ӧ�ú���
'-------------------------------------------------------------------------------------------------------

Public Function XWShowImage(ByVal lngViewerType As Long, ByVal strFilter As String, Optional ByVal lngPlanId As Long = 1) As Long
''--------------------------------------------
''���ܣ� ��������ADViewer����WEB Viewer
''������    lngViewerType -- ��Viewer�ķ�ʽ��1-�����ADViewer��2-�ٴ�WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    On Error GoTo err
    
    '��¼�ӿ���־
    If gblnXWLog = True Then
        Call WriteCommLog("XWShowImage", "XW�ӿ�", "����ADViewer����WEB��Ƭ����Ƭ��ʽ�ǣ� " & IIf(lngViewerType = 1, "ADViewer", "WEB"))
    End If
    
    If lngViewerType = 1 Then
        Call XWADViewerOpen(strFilter, lngPlanId)
    ElseIf lngViewerType = 2 Then
        Call XWWebViewerOpen(strFilter)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWWebViewerOpen(ByVal lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    
    On Error GoTo err
    
    strPath = zlDatabase.GetPara("XWWEB��Ƭ��ַ", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    If strPath <> "" Then
        strPath = Replace(strPath, "[@STU_NO]", lngOrderID)
        strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        
        '��¼�ӿ���־
        If gblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW�ӿ�", "ͨ��WEB��ʽ��Ƭ�� " & strURL)
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        '��¼�ӿ���־
        If gblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW�ӿ�", "ͨ��WEB��ʽ��Ƭ��WEB������IP��ַΪ�ա�")
        End If
        
        MsgBox "WEB������IP��ַΪ�գ��������ú�WEB��������", vbOKOnly, "��ʾ��Ϣ"
        
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWShowUnMatched(frmParent As Form, lngOrderID As Long, strModality As String) As Long
''--------------------------------------------
''���ܣ� ��������SQLServer���ݿ⣬����V_OEM_STUDY_UNMATCHED��ͼ����ʾδƥ��ļ�¼
''������    frmParent -- ������
'           lngOrderID -- ��Ҫ��������ҽ��ID
'           strModality --- Ӱ�����
''���أ�0-�ɹ�;1-δ����;2-����
''--------------------------------------------
    Dim lngResult As Long    '���������ݿ��ж�ȡ�����ļ������
    Dim strSQL As String
    Dim rsOrderInfo As ADODB.Recordset
    Dim strStudyDate As String
    Dim blnOpenDb As Boolean
    
    On Error GoTo err
    
    XWShowUnMatched = 1
    
    '��ʾδƥ��ļ�¼
    lngResult = frmXWRelateImage.zlShowMe(frmParent, lngOrderID, True, strModality)
        
    Select Case lngResult
        Case 1
            XWShowUnMatched = 1
        Case 2
            XWShowUnMatched = 2
        Case Else
            XWShowUnMatched = 0
    End Select
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    
    XWShowUnMatched = 2
End Function

Public Function XWShowMatched(frmParent As Form, lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������SQLServer���ݿ⣬��V_OEM_SERIES��ͼ����ʾ��ƥ��ļ�¼
''������    frmParent -- ������
'           lngOrderID -- ��Ҫ��������ҽ��ID
''���أ�0-�ɹ�;1-δȡ��������2-����
''--------------------------------------------
    Dim lngResult As Long    '���������ݿ��ж�ȡ�����ļ������
    
    On Error GoTo err
    
    XWShowMatched = 1
    
    '��ʾ��ƥ��ļ�¼
    lngResult = frmXWRelateImage.zlShowMe(frmParent, lngOrderID, False, "")
    
    Select Case lngResult
        Case 1
            XWShowMatched = 1
        Case 2
            XWShowMatched = 2
        Case Else
            XWShowMatched = 0
    End Select
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    
    XWShowMatched = 2
End Function

Public Function XWShowFilmPrintWind(ByVal strOrderNo As String, objOwner As Object) As Boolean
'��ʾ��Ƭ��ӡ����
    Dim objFilmPrintWind As New frmFilm
    
    XWShowFilmPrintWind = objFilmPrintWind.ShowFilmPrintWnd(strOrderNo, objOwner)
End Function


Public Function XWUnmatchSeries(ByVal lngStudyId As Long, ByVal strSeriesIds As String) As Long
'ȡ�����й���
''���أ�0-�ɹ�;1-δȡ��������2-����
''--------------------------------------------
    Dim blnOpenDb As Boolean
    Dim strSQL As String

    XWUnmatchSeries = 1
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        blnOpenDb = XWDBServerOpen
    End If
    
    '������ݿ���δ�򿪳ɹ������˳�����
    If gcnXWDBServer.State <> adStateOpen Then
        MsgBox "PACS���ݿ�������������ӣ��ò��������ܼ�����", vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    '���������洢���̡�P_OEM_UNMATCHING_RIS����ȡ������
    strSQL = "p_Oem_Split_Study(" & lngStudyId & ",'" & strSeriesIds & "')"
    gcnXWDBServer.Execute strSQL
    
    
    '������ڹ����д򿪵����ݿ����ӣ����˳�ʱ�ر�����
    If blnOpenDb = True Then
        Call XWDBServerClose
    End If
    
    XWUnmatchSeries = 0
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    XWUnmatchSeries = 2
End Function

Public Function XWUnmatchImage(lngOrderID As Long, lngXWStudyID As Long) As Long
''--------------------------------------------
''���ܣ� ��������SQLServer���ݿ⣬����P_OEM_UNMATCHING_RIS����ȡ��ָ����¼�Ĺ���
''������    lngOrderID -- ��Ҫȡ����������ҽ��ID
''          lngXWStudyID -- ��Ҫȡ�����������������ţ�0��ʾɾ��ҽ��ID�µ����м��
''���أ�0-�ɹ�;1-δȡ��������2-����
''--------------------------------------------
    Dim blnOpenDb As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    XWUnmatchImage = 1
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        blnOpenDb = XWDBServerOpen
    End If
    
    '������ݿ���δ�򿪳ɹ������˳�����
    If gcnXWDBServer.State <> adStateOpen Then
        MsgBox "PACS���ݿ�������������ӣ��ò��������ܼ�����", vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    
    '���lngXWStudyID=0������Ҫ���������ݿ��в���ҽ��ID��Ӧ�ļ���
    If lngXWStudyID = 0 Then
        strSQL = "select distinct F_STU_ID as Study���� from V_OEM_SERIES where F_STU_NO ='" & lngOrderID & "'"
        Set rsTemp = gcnXWDBServer.Execute(strSQL)
        If Not rsTemp.EOF Then
            lngXWStudyID = rsTemp!Study����
            rsTemp.MoveNext
        End If
    End If

    While lngXWStudyID <> 0
    
        '���������洢���̡�P_OEM_UNMATCHING_RIS����ȡ������
        strSQL = "P_OEM_UNMATCHING_RIS(" & lngXWStudyID & ")"
        gcnXWDBServer.Execute strSQL
        
        If rsTemp Is Nothing Then
            lngXWStudyID = 0
        Else
            If Not rsTemp.EOF Then
                lngXWStudyID = rsTemp!Study����
                rsTemp.MoveNext
            Else
                lngXWStudyID = 0
            End If
        End If
    Wend
    
    '���������洢���̣�ȡ������
    strSQL = "select F_SER_ID as SERIES����,F_STU_ID as Study����,F_SER_UID as ����UID,F_SER_DATE as ��������,F_SER_TIME as ����ʱ��, " _
            & " F_SER_CONTEXT as ��������,F_MODALITY as Ӱ������,F_STU_NO as ҽ��ID from V_OEM_SERIES where F_STU_NO ='" & lngOrderID _
            & "' order by F_STU_ID ,F_SER_ID"
    Set rsTemp = gcnXWDBServer.Execute(strSQL)
    If rsTemp.EOF = True Then
        strSQL = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsUnmatchImage(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure strSQL, "ȡ������"
    End If
    
    '������ڹ����д򿪵����ݿ����ӣ����˳�ʱ�ر�����
    If blnOpenDb = True Then
        Call XWDBServerClose
    End If
    
    XWUnmatchImage = 0
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    XWUnmatchImage = 2
End Function


'---------------------------------------------------------------------------------------
'���²�����Ϣ
'---------------------------------------------------------------------------------------
Public Function XWStudyUpdate(ByVal lngAdviceID As Long, _
    ByVal strName As String, ByVal strSex As String, ByVal strAge As String) As Long
    Dim strSQL As String
    Dim blnOpenDb As Boolean
    
    
    blnOpenDb = False
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            blnOpenDb = True
        End If
    End If
    
    strSQL = "P_OEM_MATCHING_RIS_SIMPLE('" & lngAdviceID & "','" & strName & "','" & strSex & "','" & strAge & "')"
    
    Call gcnXWDBServer.Execute(strSQL)
    
    
    '������ڹ����д򿪵����ݿ����ӣ����˳�ʱ�ر�����
    If blnOpenDb = True Then
        Call XWDBServerClose
    End If
End Function


'---------------------------------------------------------------------------------------
'���ձ����Windows��Ϣ
'---------------------------------------------------------------------------------------

Public Function XWHook(ByVal hWnd As Long) As Long
    'ָ���Զ���Ĵ��ڹ���
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��
    If App.LogMode <> 0 Then
        Call WriteCommLog("XWHook", "XW�ӿ�", "������Ϣ������̡�")
        
        XWHook = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf XWWindowProc)
        Debug.Print "Hooked"
    End If
End Function

Public Sub XWUnhook(ByVal hWnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
        Call WriteCommLog("XWHook", "XW�ӿ�", "ж����Ϣ������̡�")
        
        temp = SetWindowLong(hWnd, GWL_WNDPROC, lpWndProc)
    End If
End Sub

Function XWWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��Ϣ�������ר�Ŵ����ض��� WM_XWReportImage ��Ϣ
    Dim strLog As String
    
    If uMsg = WM_XWReportImage Then
        strLog = Now & " umsg = " & uMsg & ";wparam = " & wParam & ";lparam = " & lParam & vbCrLf
        
        If gblnXWLog = True Then
            Call WriteCommLog("XWWindowProc", "XW�ӿ�", strLog)
        End If
        '�����������͵�ϵͳ������ı���ͼ��
        If lParam <> 0 Then
            If gblnXWLog Then
                Call WriteCommLog("XWWindowProc", "XW�ӿ�", "���뱨��ͼ������̡�")
            End If
            
            Call XWSaveReportImages(lParam)
        End If
    End If
  
    '����ԭ���Ĵ��ڹ���
    XWWindowProc = CallWindowProc(plngXWPreWndProc, hw, uMsg, wParam, lParam)
End Function

Public Sub XWSaveReportImages(lngOrderID As Long)
'------------------------------------------------
'���ܣ���ͼ��Ӽ����屣��ɱ���ͼ
'������ lngOrderID -- ҽ��ID
'���أ�
'------------------------------------------------
    Dim dcmImage As New DicomImage
    Dim strFileName As String
    Dim strLocalPath As String
    Dim dcmG As New DicomGlobal
    Dim strTempPath As String, lngBuffSize As Long
    Dim strStudyUID As String
    Dim strDeviceNO As String
    Dim strFTPIP As String
    Dim strFtpUrl As String
    Dim strFtpVirtualPath As String
    Dim strFTPUser As String
    Dim strFTPPwd As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim Inet As New clsFtp
    Dim lngResult As Long

    On Error GoTo err

    If gfrmPacsMain Is Nothing Then Exit Sub

    '�Ӽ������ȡ����ͼ
    dcmImage.Paste
    '���ݹ����������ͼ����
    dcmG.RegString("UIDRoot") = "1"
    strFileName = dcmG.NewUID & ".jpg"
    
    '��ȡ�洢�豸������FTP�еĸ���Ŀ¼
    strSQL = "select ���UID FROM Ӱ�����¼ where ҽ��ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���UID", lngOrderID)
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "��ѯ���UID����ѯ��SQL�ǣ�" & strSQL & "  ҽ��ID=" & lngOrderID & vbCrLf & "��ѯ���ļ�¼��Ϊ�� " & rsTemp.RecordCount)
    End If
    
    If rsTemp.EOF = True Then Exit Sub
    
    strStudyUID = rsTemp!���uid
    Call GetDeptStorageDevice(gfrmPacsMain, lngOrderID, strStudyUID, glngXWDeptID, G_LNG_PACSSTATION_MODULE, gblnXWMoved, strDeviceNO, _
                strFTPIP, strFtpUrl, strFtpVirtualPath, strFTPUser, strFTPPwd)

    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "��ȡ�洢�豸�����UID = " & strStudyUID & "��ҽ��ID = " & lngOrderID & "��ִ�п���ID= " & glngXWDeptID & "���洢�豸��= " & strDeviceNO)
    End If
    
    '��ȡ������ʱ�ļ���
    strFtpVirtualPath = Replace(strFtpVirtualPath, strFtpUrl, "")
    strLocalPath = App.Path & "\TmpImage\" & Replace(strFtpVirtualPath, "/", "\")
    '��������Ŀ¼
    Call MkLocalDir(strLocalPath)

    '������ͼ������ļ�
    dcmImage.FileExport strLocalPath & "\" & strFileName, "JPG"
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "���汾��ͼ��ͼ���ļ���Ϊ��" & strLocalPath & "\" & strFileName)
    End If

    '������ͼ�ϴ���FTPĿ¼,�����浽���ݿ�
    lngResult = Inet.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    If lngResult <> 0 Then
        lngResult = Inet.FuncUploadFile(strFtpUrl & strFtpVirtualPath, strLocalPath & "\" & strFileName, strFileName)
        Inet.FuncFtpDisConnect
        
        If gblnXWLog = True Then
            Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "�ϴ�FTPͼ��FTP IP��ַ= " & strFTPIP & "��FTP��Ŀ¼=" & strFtpUrl & strFtpVirtualPath & "��ͼ���ļ���Ϊ��" & strFileName)
        End If
        
         '�޸����ݿ⣬���ӱ���ͼ
        If lngResult = 0 Then
            strSQL = "ZL_Ӱ���鱨��_ADD('" & strStudyUID & "','" & strFileName & "')"
            zlDatabase.ExecuteProcedure strSQL, "���汨��ͼ��"
            
            If gblnXWLog = True Then
                Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "���ݿⱣ�汨��ͼ��ִ�д洢���̣�" & strSQL)
            End If
        
            strSQL = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsSetFTPDeviceNo(" & lngOrderID & ",'" & strDeviceNO & "')"
            zlDatabase.ExecuteProcedure strSQL, "���汨��ͼ���FTP�豸��"
            
            If gblnXWLog = True Then
                Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "���汨��ͼ���豸�ţ�ִ�д洢���̣�" & strSQL)
            End If
        End If
    End If

    Exit Sub
err:
    '����������ʾ
    Debug.Print Now & "����:" & err.Description & vbCrLf
    Inet.FuncFtpDisConnect
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "���汨��ͼ�������������ǣ�" & err.Description)
    End If
        
End Sub

Public Sub subXWShowArchiveManager(intType As Integer)
'------------------------------------------------
'���ܣ���������ArchiveManagerʵ�ֶ���Ĺ���
'������ intType = 1---ɾ��ͼ��2---����ͼ��3---���̿�¼
'���أ��ޣ�ֱ�Ӵ�ArchiveManager
'------------------------------------------------
    Dim strCommand As String
    Dim strUser As String
    Dim strPswd As String
    
    On Error GoTo err
    
    If intType = 1 Then     'ɾ��ͼ��
        strUser = zlDatabase.GetPara("XWɾ��ͼ���û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
        strPswd = zlDatabase.GetPara("XWɾ��ͼ������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    ElseIf intType = 2 Then     '����ͼ��
        strUser = zlDatabase.GetPara("XW����ͼ���û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
        strPswd = zlDatabase.GetPara("XW����ͼ������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    Else    '���̿�¼
        strUser = zlDatabase.GetPara("XW���̿�¼�û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
        strPswd = zlDatabase.GetPara("XW���̿�¼����", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    End If
    
    
    If strUser <> "" And strPswd <> "" Then
        strCommand = "C:\PACS\ArchiveManager.exe " & strUser & "^" & strPswd
        
        If gblnXWLog = True Then
            Call WriteCommLog("subXWShowArchiveManager", "XW�ӿ�", "��ArchiveManager������" & IIf(intType = 1, "ɾ��ͼ��", IIf(intType = 2, "����ͼ��", "���̿�¼")) & "�Ĳ����������ǣ�" & strCommand)
        End If
            
        Shell strCommand, vbMaximizedFocus
    Else
        If gblnXWLog = True Then
            Call WriteCommLog("subXWShowArchiveManager", "XW�ӿ�", "����" & IIf(intType = 1, "ɾ��ͼ��", IIf(intType = 2, "����ͼ��", "���̿�¼")) & "�Ĳ���ʱ���û��������벻��Ϊ�ա� " & vbCrLf _
                & "�û����ǣ�" & strUser & "�������ǣ�" & strPswd)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub
