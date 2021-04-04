Attribute VB_Name = "mdlXWPACS"
Option Explicit

Public gcnXWDBServer As New ADODB.Connection         '�������ݿ�����
Public gfrmPacsMain As frmPacsMain                  '�������ձ���ͼ��Ϣ�Ĵ���ָ��
Public glngXWDeptID As Long                         '��ǰ����ID
Public gblnXWMoved As Boolean                       '�Ƿ�ת��
Public gblnXWLog As Boolean                         '�Ƿ��¼ͨѶ��־
Public gstrOracleOwner As String                    'oracle����ӵ����
Public gblnUseXinWangPacs As Boolean                '�ж��Ƿ�������������Ƭ
Public gstrImageShareDir As String                  '�ϰ��Ӱ����洢Ŀ¼

'���������ṩ�� InterCOM.dll������ADViewer�鿴ͼ��

'��������
Public Declare Function OEMViewStart Lib "InterCOM.dll" (ByVal cpReserved1 As String, ByVal cpReserved2 As String, ByVal cpReserved3 As String) As Long
Public Declare Function OEMViewExit Lib "InterCOM.dll" (ByVal cpReserved1 As String, ByVal cpReserved2 As String, ByVal cpReserved3 As String) As Long
Public Declare Function OEMViewOpen Lib "InterCOM.dll" (ByVal lPlanID As Long, ByVal cpFilter As String, ByVal lFunc As Long, ByVal cpReserved As String) As Long
Public Declare Function OEMViewClose Lib "InterCOM.dll" (ByVal cpReserved As String) As Long
 
'�����ṩ�ĺ���
'����ADViewer�� long OEMViewStart ( LPCTSTR cpReserved1, LPCTSTR cpReserved2, LPCTSTR cpReserved3 );
'�˳�ADViewer�� long OEMViewExit ( LPCTSTR cpReserved1, LPCTSTR cpReserved2, LPCTSTR cpReserved3 );
'��ָ��ͼ�� long OEMViewOpen ( long lPlanID, LPCTSTR cpFilter, long lFunc, LPCTSTR cpReserved );
'�ر�ͼ�� long OEMViewClose ( LPCTSTR cpReserved );


'���ձ���ͼ����Ϣ��API
Public Const WM_XWReportImage As Long = 5120
'��ϢHook����
Public plngXWPreWndProc As Long       'ԭ������Ϣ�������


'-----------------------------------------------------------------------------------------------------
'ADViewer��������
'-----------------------------------------------------------------------------------------------------

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
    Dim strSql As String
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

Public Function XWDBServerOpen() As Long
'--------------------------------------------
'���ܣ� ��������SQLServer���ݿ�
'��������
'���أ�0-�ɹ�
'--------------------------------------------
    Dim strSqlUser As String
    Dim strSqlPWD As String
    Dim strDataSource As String

    gblnUseXinWangPacs = False
    
    If InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "����") <= 0 Then Exit Function
    
    '������ORACLE ģ������л�ȡ���������ݿ������IP��ַ���û���������
    strDataSource = zlDatabase.GetPara("XW���ݿ������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    strSqlUser = zlDatabase.GetPara("XW���ݿ�������û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    strSqlPWD = zlDatabase.GetPara("XW���ݿ����������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    gstrImageShareDir = zlDatabase.GetPara("XW��ʷͼ����Ŀ¼", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    
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
    
    gblnUseXinWangPacs = True
    
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
    Dim strIP As String
    Dim strURL As String
    
    On Error GoTo err
    
    strIP = zlDatabase.GetPara("XWWEB������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    If strIP <> "" Then
        strURL = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIP & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
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
    Dim lngXWStudyID As Long    '���������ݿ��ж�ȡ�����ļ������
    Dim strSql As String
    Dim rsOrderInfo As ADODB.Recordset
    Dim strStudyDate As String
    Dim blnOpenDB As Boolean
    
    On Error GoTo err
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            blnOpenDB = True
        End If
    End If
    
    XWShowUnMatched = 2
    
    '��ʾδƥ��ļ�¼
    lngXWStudyID = frmXWRelateImage.zlShowMe(frmParent, lngOrderID, True, strModality)
    If lngXWStudyID > 0 Then
        strStudyDate = frmXWRelateImage.pstrStudyDate
        'ʹ�����ҽ��ID����ͼ�����
        
        strSql = "Select b.����ID,b.�����,b.סԺ��,b.������ as ����,b.����,b.�Ա�,b.����,To_char(b.��������,'yyyymmdd') As ��������, " _
                    & " c.Ӣ���� as ƴ����,c.Ӱ�����,c.����,a.������Դ,a.ִ�п���ID,d.���� As ִ�п���,a.����ʱ��,a.��ʼִ��ʱ�� " _
                    & " From ����ҽ����¼ a,������Ϣ b,Ӱ�����¼ c,���ű� d  " _
                    & " Where a.����Id = b.����ID And a.Id = c.ҽ��ID And a.ִ�п���ID =d.Id  and a.Id = [1]"
        Set rsOrderInfo = zlDatabase.OpenSQLRecord(strSql, "��ѯ�����Ϣ", lngOrderID)
        
        If rsOrderInfo.RecordCount <> 0 Then
            '���������洢���̡�P_OEM_MATCHING_RIS��������ͼ��
                    
            strSql = "P_OEM_MATCHING_RIS(" & lngXWStudyID & ",'" & lngOrderID & "','" & rsOrderInfo!����ID & "','" & Nvl(rsOrderInfo!�����, 0) _
                    & "','" & Nvl(rsOrderInfo!סԺ��, 0) & "','" & Nvl(rsOrderInfo!����, 0) & "','" & Nvl(rsOrderInfo!����) & "','" _
                    & Nvl(rsOrderInfo!�Ա�) & "','" & Nvl(rsOrderInfo!����, 0) & "','" & Nvl(rsOrderInfo!��������) & "','" & Nvl(rsOrderInfo!ƴ����) _
                    & "','" & Nvl(rsOrderInfo!Ӱ�����) & "'," & rsOrderInfo!���� & "," & Nvl(rsOrderInfo!������Դ, 3) & "," & Nvl(rsOrderInfo!ִ�п���ID) _
                    & ",'" & Nvl(rsOrderInfo!ִ�п���) & "','','')"
                    
            gcnXWDBServer.Execute strSql
            
            '���������洢����"b_XINWANGInterface.PacsStatusChange"������ͼ��
            strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsStatusChange(1," & lngOrderID & ",'" & Nvl(rsOrderInfo!Ӱ�����) & "'," & rsOrderInfo!���� & ",to_date('" _
                        & Trim(strStudyDate) & "','YYYY.MM.DD'),null,null)"
            zlDatabase.ExecuteProcedure strSql, "����ͼ��"
        End If
        XWShowUnMatched = 0
        
    ElseIf lngXWStudyID = -1 Then
        'ͼ������޸�
        XWShowUnMatched = 0
        
    Else
        XWShowUnMatched = 1
    End If
    
    '������ڹ����д򿪵����ݿ����ӣ����˳�ʱ�ر�����
    If blnOpenDB = True Then
        Call XWDBServerClose
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function XWShowMatched(frmParent As Form, lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������SQLServer���ݿ⣬��V_OEM_SERIES��ͼ����ʾ��ƥ��ļ�¼
''������    frmParent -- ������
'           lngOrderID -- ��Ҫ��������ҽ��ID
''���أ�0-�ɹ�;1-δȡ��������2-����
''--------------------------------------------
    Dim lngXWStudyID As Long    '���������ݿ��ж�ȡ�����ļ������
    
    On Error GoTo err
    
    XWShowMatched = 2
    
    '��ʾ��ƥ��ļ�¼
    lngXWStudyID = frmXWRelateImage.zlShowMe(frmParent, lngOrderID, False, "")
    If lngXWStudyID <> 0 Then
        'ʹ�����ҽ��IDȡ��ͼ�����
        'ʹ�����lngXWStudyID���������ݿ���ȡ������
        
        If MsgBoxD(frmParent, "�Ƿ�ȷ��ȡ��ͼ��ͼ����Ϣ�Ĺ�����", vbOKCancel, "��ʾ��Ϣ") = vbCancel Then
            Exit Function
        End If
        
        XWShowMatched = XWUnmatchImage(lngOrderID, lngXWStudyID)
    Else
        XWShowMatched = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function XWUnmatchImage(lngOrderID As Long, lngXWStudyID As Long) As Long
''--------------------------------------------
''���ܣ� ��������SQLServer���ݿ⣬����P_OEM_UNMATCHING_RIS����ȡ��ָ����¼�Ĺ���
''������    lngOrderID -- ��Ҫȡ����������ҽ��ID
''          lngXWStudyID -- ��Ҫȡ�����������������ţ�0��ʾɾ��ҽ��ID�µ����м��
''���أ�0-�ɹ�;1-δȡ��������2-����
''--------------------------------------------
    Dim blnOpenDB As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset

    XWUnmatchImage = 1
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            blnOpenDB = True
        End If
    End If
    
    '���lngXWStudyID=0������Ҫ���������ݿ��в���ҽ��ID��Ӧ�ļ���
    If lngXWStudyID = 0 Then
        strSql = "select distinct F_STU_ID as Study���� from V_OEM_SERIES where F_STU_NO ='" & lngOrderID & "'"
        Set rsTemp = gcnXWDBServer.Execute(strSql)
        If Not rsTemp.EOF Then
            lngXWStudyID = rsTemp!Study����
            rsTemp.MoveNext
        End If
    End If

    While lngXWStudyID <> 0
    
        '���������洢���̡�P_OEM_UNMATCHING_RIS����ȡ������
        strSql = "P_OEM_UNMATCHING_RIS(" & lngXWStudyID & ")"
        gcnXWDBServer.Execute strSql
        
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
    strSql = "select F_SER_ID as SERIES����,F_STU_ID as Study����,F_SER_UID as ����UID,F_SER_DATE as ��������,F_SER_TIME as ����ʱ��, " _
            & " F_SER_CONTEXT as ��������,F_MODALITY as Ӱ������,F_STU_NO as ҽ��ID from V_OEM_SERIES where F_STU_NO ='" & lngOrderID _
            & "' order by F_STU_ID ,F_SER_ID"
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    If rsTemp.EOF = True Then
        strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsUnmatchImage(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure strSql, "ȡ������"
    End If
    
    '������ڹ����д򿪵����ݿ����ӣ����˳�ʱ�ر�����
    If blnOpenDB = True Then
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
'���ձ����Windows��Ϣ
'---------------------------------------------------------------------------------------

Public Function XWHook(ByVal hWnd As Long) As Long
    'ָ���Զ���Ĵ��ڹ���
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��
    If App.LogMode <> 0 Then
        XWHook = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf XWWindowProc)
        Debug.Print "Hooked"
    End If
End Function

Public Sub XWUnhook(ByVal hWnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
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
    Dim strFtpIp As String
    Dim strFtpUrl As String
    Dim strFtpVirtualPath As String
    Dim strFTPUser As String
    Dim strFTPPwd As String
    Dim strSql As String
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
    strSql = "select ���UID FROM Ӱ�����¼ where ҽ��ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ���UID", lngOrderID)
    
    If gblnXWLog = True Then
        Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "��ѯ���UID����ѯ��SQL�ǣ�" & strSql & vbCrLf & "��ѯ���ļ�¼��Ϊ�� " & rsTemp.RecordCount)
    End If
    
    If rsTemp.EOF = True Then Exit Sub
    
    strStudyUID = rsTemp!���UID
    Call GetDeptStorageDevice(gfrmPacsMain, lngOrderID, strStudyUID, glngXWDeptID, G_LNG_PACSSTATION_MODULE, gblnXWMoved, strDeviceNO, _
                strFtpIp, strFtpUrl, strFtpVirtualPath, strFTPUser, strFTPPwd)

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
    lngResult = Inet.FuncFtpConnect(strFtpIp, strFTPUser, strFTPPwd)
    If lngResult <> 0 Then
        lngResult = Inet.FuncUploadFile(strFtpUrl & strFtpVirtualPath, strLocalPath & "\" & strFileName, strFileName)
        Inet.FuncFtpDisConnect
        
        If gblnXWLog = True Then
            Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "�ϴ�FTPͼ��FTP IP��ַ= " & strFtpIp & "��FTP��Ŀ¼=" & strFtpUrl & strFtpVirtualPath & "��ͼ���ļ���Ϊ��" & strFileName)
        End If
        
         '�޸����ݿ⣬���ӱ���ͼ
        If lngResult = 0 Then
            strSql = "ZL_Ӱ���鱨��_ADD('" & strStudyUID & "','" & strFileName & "')"
            zlDatabase.ExecuteProcedure strSql, "���汨��ͼ��"
            
            If gblnXWLog = True Then
                Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "���ݿⱣ�汨��ͼ��ִ�д洢���̣�" & strSql)
            End If
        
            strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsSetFTPDeviceNo(" & lngOrderID & ",'" & strDeviceNO & "')"
            zlDatabase.ExecuteProcedure strSql, "���汨��ͼ���FTP�豸��"
            
            If gblnXWLog = True Then
                Call WriteCommLog("XWSaveReportImages", "XW�ӿ�", "���汨��ͼ���豸�ţ�ִ�д洢���̣�" & strSql)
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
