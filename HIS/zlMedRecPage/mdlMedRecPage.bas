Attribute VB_Name = "mdlMedRecPage"
Option Explicit
'-----------------------------------------------------------
'��׼���벡��ϵͳ������Ŀ
'-----------------------------------------------------------
'1------�ӿڱ���
Public gcnOracle As ADODB.Connection    '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gclsMipModule As zl9ComLib.clsMipModule
'2------ȫ�ֱ���
Public gclsMain As Object                         '��ǰ����ʵ��
Public gclsPros As clsProperty                   '������
Public gstrSysName As String                       'ϵͳ����
Public gstrProductName As String               'OEM��Ʒ����
Public gstrUnitName As String                     '�û���λ����
Public gobjReport As clsReport                    '�����ӡ������������ҳ��ӡԤ��
Public gcolPrivs As Collection                      '��¼�ڲ�ģ���Ȩ��
Public UserInfo As TYPE_USER_INFO
Public gobjPatient As Object                        '������Ϣ�ӿ�
'5------ȫ�ֱ���
Public grsDeptInfo As ADODB.Recordset         '�ٴ����һ����¼��
Public gintCA As Integer                           '����ǩ����֤����
Public gstrESign As String                         '����ǩ�����Ƴ���
Public grsSign As Recordset                      '����ǩ�����ò���
Public gobjRis As Object                            '����RIS�ӿ�
Public gblnSet  As Boolean                         '�м������ֹ�¼��ظ�����
Public gColErr As New Collection
Public gColWarn As New Collection
Public gColCtl As New Collection                    '��ҳ�ؼ�����

Public gBlnNew As Boolean                     '�Ƿ�����ҳ��Ҹ�ҳ
Public gfrmMecCol As Collection                     '��Ҹ�ҳ����
Public gPic��Ҹ�ҳ As Integer             '��Ҹ�ҳpicturebox��index
Public colErrTmp As Collection                 '��Ҳ�����ʾ��Ϣ����
Public gIntPic As Integer                  'Դ�����picturebox������
'-----------------------------------------------------------
'��׼��(����)
'-----------------------------------------------------------
'1------�ӿڱ���
Public gblnHaveOPS As Boolean               '�Ƿ�װ����ϵͳ��ϵͳ��= 2400
Public gobjCommunity As Object              '���������ӿڶ���
Public gobjPass As Object                          '������ҩ�ӿڶ���
Public gobjESign As Object                        'ǩ����������
Public gclsInsure As New clsInsure           'ҽ������
Public gobjPlugIn As Object                      '��ҹ��ܶ���
'-----------------------------------------------------------
''����ϵͳ(����)
'-----------------------------------------------------------
'5------ȫ�ֱ���
Public grsBabyInfo As ADODB.Recordset         '�и��������������Ϣ
Public grsBabyDiag As ADODB.Recordset        '�и�������������������Ϣ
Public grsDeliceryInfo As ADODB.Recordset    '�и�����������Ϣ
'��־���ٶ���
Public gobjLog As TextStream
Public gobjFSO As New FileSystemObject
Public gblnUnload As Boolean '���ڼ�¼��ҳ�����Ƿ��Ѿ�ж��

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

Public Function CheckShare(ByVal lngSysShareNO As Long, Optional ByVal lngSysMainNO As Long = 100) As Boolean
'���ܣ���׼ϵͳ������ϵͳ�Ƿ��ǹ���װ
'������lngSysShareNO= ����װ��ϵͳ
'           lngSysMainNO=��ϵͳ
    Dim lngShareNum As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
'Select * From (Select * From zlSystems Start With ��� = 100 Connect By Prior ��� = �����) Where ��� = 300
'Select * From (Select * From zlSystems Start With ��� = 300 Connect By Prior ��� = �����) Where ��� = 100
    strSQL = "Select s.���" & vbNewLine & _
            "From zlSystems S" & vbNewLine & _
            "Where s.������װ = 1 And s.���  = [1] And s.����� = [2]"
    On Error GoTo errH
    '���ڴ��ڶ������������׼������ױ��100��101������������199����������ж�
    '�����ײ��ܹ���װ
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, (lngSysShareNO \ 100) * 100, (lngSysMainNO \ 100) * 100)
    CheckShare = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckHaveSys(ByVal lngSysNO As Long) As Boolean
'���ܣ��ж��Ƿ�װ��ĳ��ϵͳ
'������lngSysNO �ж���ϵͳ
    Dim lngShareNum As Long
    Dim strSQL As String
    Dim rsTmp As Recordset

    strSQL = "Select s.���" & vbNewLine & _
            "From zlSystems S" & vbNewLine & _
            "Where s.������װ = 1 And Floor(s.��� / 100) = [1] "

    On Error GoTo errH
    '���ڴ��ڶ������������׼������ױ��100��101������������199����������ж�
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngSysNO \ 100)

    If rsTmp.RecordCount > 0 Then CheckHaveSys = True: Exit Function
    CheckHaveSys = False
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
'���ܣ�ת����ֵΪ����
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String

    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)

    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"

    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay

    If Not IsDate(strDate) Then Exit Function

    strDate = Format(strDate, "yyyy-mm-dd")
    If blnDec Then strDate = DateAdd("d", -1, Format(strDate, "yyyy-mm-dd"))
    TranNumToDate = strDate
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function HaveRIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� �����ӿڲ��� �Ƿ����
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If gobjRis Is Nothing Then
        On Error Resume Next
        Set gobjRis = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
    End If
    If gobjRis Is Nothing Then
        If blnMsg Then
            MsgBox "�����ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    HaveRIS = True
End Function
