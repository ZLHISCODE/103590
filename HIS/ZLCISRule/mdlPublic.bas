Attribute VB_Name = "mdlPublic"

Option Explicit
'API
'������ָ������Ļ�����ϵ�λ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'��ô�������Ļ�����е�λ��
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'�ж�ָ���ĵ��Ƿ���ָ���ľ����ڲ�
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'׼������ʹ����ʼ������ǰ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'�����ƶ�����
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'��ȡ����״̬
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'HWND hwnd, // ָ���ֲ㴰�ھ��
'COLORREF crKey, // ָ����Ҫ͸���ı�����ɫֵ������RGB()��
'BYTE bAlpha, // ����͸���ȣ�0��ʾ��ȫ͸����255��ʾ��͸��
'DWORD dwFlags // ͸����ʽ
'       ���У�dwFlags������ȡ����ֵ��
'       LWA_ALPHA=&H2ʱ��crKey������Ч��bAlpha������Ч��
'       LWA_COLORKEY=&H1�������е�������ɫΪcrKey�ĵط�����Ϊ͸����bAlpha������Ч�䳣��ֵΪ1
'       LWA_ALPHA | LWA_COLORKEY��crKey�ĵط�����Ϊȫ͸�����������ط�����bAlpha����ȷ��͸����
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'�Զ�����������
Public Type TYPE_USER_INFO
    id As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    רҵ�������� As String
    ��ҩ���� As Long
End Type

Public UserInfo As TYPE_USER_INFO

'����
Public Const GWL_WNDPROC = -4&
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const SWP_NOACTIVATE = &H10 '�������
Public Const GWL_EXSTYLE  As Long = (-20)
Public Const WS_EX_TOPMOST As Long = &H8
Public Const HWND_TOPMOST As Long = -1
Public Const SW_SHOWMAXIMIZED = 3
'API:GetSystemMetrics
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Const conCOLOR_TITLE_BAR As Long = 16298544 '16298544 rgb(48,178,248); 14392064 'RGB(0, 155, 219)

'���б���
Public glngOldWindowProc As Long '��������ϵͳĬ�ϵĴ�����Ϣ�������ĵ�ַ
Public glngSys As Long

Public gobjLIS As Object    '

Public gstrSysName As String                'ϵͳ����
Public gstrProductName As String
Public gstrUnitName As String
Public gblnLog As Boolean         'T-������־����;F-�ر���־����
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��


Public Function GetXMLResult(ByVal rsRec As ADODB.Recordset)
'����:���췴��������ӦXML
    Dim i As Long
    Dim strXML As String
    For i = 1 To rsRec.RecordCount
        strXML = strXML & "    <info name=""" & rsRec!Name & """ type=""" & rsRec!Type & """ index=""" & _
            rsRec!index & """ value=""" & rsRec!Default & """ obsid=""" & rsRec!Obsid & """/>" & vbNewLine
        rsRec.MoveNext
    Next
    GetXMLResult = Replace(strXML, """", "\""")
End Function


'�Զ������Ϣ������
Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'����:��������¼����д���,�ǹ����¼�����Ĭ�ϴ�����Ϣ������
'����:vsc-VScrollBar ����
'     OldWindowProc Ĭ�ϴ�����Ϣ��������ַ
    On Error Resume Next
    If msg = WM_MOUSEWHEEL Then
        '���������¼����д���
        If wParam = -7864320 Then '���¹���
            If frmInquiryInfo.vsc.Value - 10 < frmInquiryInfo.vsc.Max Then
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Max
            Else
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Value - 10
            End If
        ElseIf wParam = 7864320 Then '���Ϲ���
            If frmInquiryInfo.vsc.Value + 10 > frmInquiryInfo.vsc.Min Then
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Min
            Else
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Value + 10
            End If
        End If
    Else
        '����Ĭ�ϴ�����Ϣ������
        NewWindowProc = CallWindowProc(glngOldWindowProc, hWnd, msg, wParam, lParam)
    End If
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.����ID = NVL(rsTmp!����ID, 0)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.רҵ����ְ�� = NVL(rsTmp!רҵ����ְ��)
            GetUserInfo = True
        End If
    End If
End Function


Public Function SubmitMainInfo(ByVal strJsonIn As String, ByRef strJsonOut As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : SubmitMainInfo
' Author    : YWJ
' Date      : 2019-09-19 16:40:39
' Parameter : strJsonIn-����JSON�ַ��� �����ʽ����:
'             strJsonOut-����JSON�ַ���
' Purpose   : �ύ������Ϣ
' Return    : T-�ɹ�;F-ʧ��
' TEST URL  : http://192.168.32.201:8889/bizdomain/07f7c460-7dd8-49b5-a79b-0a90b9369224
'---------------------------------------------------------------------------------------
    Dim strErr As String
    Dim blnRet As Boolean
    
    '--��־
    WriteLog "���⣺�ύ������Ϣ" & vbNewLine & _
             "������SubmitMainInfo" & vbNewLine & _
             "��Σ�" & strJsonIn & vbNewLine
    blnRet = Sys.NewSystemSvr("֪ʶ��", "�ύ������Ϣ", strJsonIn, strJsonOut, strErr)
    WriteLog "���⣺�ύ������Ϣ" & vbNewLine & _
             "������SubmitMainInfo" & vbNewLine & _
             "���Σ�" & strJsonOut & vbNewLine & _
             "����ֵ��" & blnRet & vbNewLine & _
             IIf(strErr <> "", "������Ϣ:" & strErr & vbNewLine, "")
    SubmitMainInfo = blnRet
     
End Function
 


Public Function SubmitInquiriyInfo(ByVal strJsonIn As String, ByRef strJsonOut As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : SubmitInquiriyInfo
' Author    : YWJ
' Date      : 2019-09-19 16:50:43
' Parameter : strJsonIn-����JSON�ַ��� �����ʽ����:
'             strJsonOut-����JSON�ַ���
' Purpose   : �ύ������Ϣ
' Return    : T-�ɹ�;F-ʧ��
' TEST URL  : http://192.168.32.201:8889/bizdomain/c6050afd-135e-454a-b53f-e5a7d7634399
'---------------------------------------------------------------------------------------
'
    Dim strErr As String
    Dim blnRet As Boolean
    
     WriteLog "���⣺�ύ������Ϣ" & vbNewLine & _
             "������SubmitInquiriyInfo" & vbNewLine & _
             "��Σ�" & strJsonIn & vbNewLine
    blnRet = Sys.NewSystemSvr("֪ʶ��", "�ύ������Ϣ", strJsonIn, strJsonOut, strErr)
    WriteLog "���⣺�ύ������Ϣ" & vbNewLine & _
             "������SubmitInquiriyInfo" & vbNewLine & _
             "���Σ�" & strJsonOut & vbNewLine & _
             "����ֵ��" & blnRet & vbNewLine & _
             IIf(strErr <> "", "������Ϣ:" & strErr & vbNewLine, "")
    SubmitInquiriyInfo = blnRet
End Function


Public Function GetPatiInfo(ByVal lngPatiID As Long, ByVal lngVisitId As Long, ByVal strRegNo As String, _
    ByVal bytScene As Byte, ByRef lngRegId As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : GetPatiInfo
' Author    : YWJ
' Date      : 2019-09-23 13:45:55
' Parameter :
'             lngPatiID -����ID
'             lngVisitId -��ҳID
'             strRegNO -�Һŵ���
'             bytScene-���� 1-����\סԺҽ���´�;2-��ϱ���;3-����;4-�걾�ɼ�
'             lngRegId-����:�Һ�ID
' Purpose   : ��ȡ������Ϣ
'---------------------------------------------------------------------------------------
'            ���˻�����Ϣ   ��ʽ����:
'            "patient_info":{
'            "pid":"5066404",
'            "visit_id":"1",
'            "visit_no":"314929",
'            "name":"������",
'            "age":"31��",
'            "birthday":"1989-10-10 09-10-10",
'            "gender":"Ů",
'            "marital_status":"�ѻ�",
'            "operator_id":"489b7bba-31cd-4f59-8fef-c12f0570db61",
'            "operator":"֣־��",
'            "enc_type":"2","scene":"1"
'            }
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strPati As String
    Dim bytType As Byte
    Dim strVisitId As String
    Dim strVisitNo As String
    On Error GoTo ErrH
    If strRegNo = "" Then
        strSQL = "Select Nvl(b.����, a.����) As ����, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, b.����״��,a.��������,b.סԺ�� as ��ʶ�� " & vbNewLine & _
                "From ������Ϣ A, ������ҳ B" & vbNewLine & _
                "Where a.����id = b.����id And b.����id = [1] And b.��ҳid = [2]"
        bytType = 2 'סԺ
        strVisitId = lngVisitId
         
    Else
                
        strSQL = "Select b.Id As ����id, Nvl(b.����, a.����) As ����, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, a.����״��, a.��������,b.����� as ��ʶ��" & vbNewLine & _
                "From ������Ϣ A, ���˹Һż�¼ B" & vbNewLine & _
                "Where a.����id = b.����id And a.����id = [1] And b.No = [3] And b.��¼���� = 1 And b.��¼״̬ = 1"
        bytType = 1 '����
        strVisitId = strRegNo
    End If

    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiInfo", lngPatiID, lngVisitId, strRegNo)
    If rsTemp.EOF Then Exit Function
    If bytType = 1 Then lngRegId = rsTemp!����ID
    strPati = "\""patient_info\"":{\""pid\"":\""" & lngPatiID & "\""," & vbNewLine & _
                                "\""visit_id\"":\""" & strVisitId & "\""," & vbNewLine & _
                                "\""visit_no\"":\""" & rsTemp!��ʶ�� & "\""," & _
                                "\""name\"":\""" & rsTemp!���� & "\""," & vbNewLine & _
                                "\""age\"":\""" & rsTemp!���� & "\""," & vbNewLine & _
                                "\""birthday\"":\""" & Format(rsTemp!��������, "YYYY-MM-DD HH:MM:SS") & "\""," & vbNewLine & _
                                "\""gender\"":\""" & rsTemp!�Ա� & "\""," & vbNewLine & _
                                "\""marital_status\"":\""" & rsTemp!����״�� & "\""," & vbNewLine & _
                                "\""operator_id\"":\""" & UserInfo.id & "\""," & vbNewLine & _
                                "\""operator\"":\""" & UserInfo.���� & "\""," & vbNewLine & _
                                "\""enc_type\"":\""" & bytType & "\""," & vbNewLine & _
                                "\""scene\"":\""" & bytScene & "\""}"
                                
    GetPatiInfo = strPati

    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog

End Function


Public Function GetMainInfo(ByVal lngPatiID As Long, ByVal lngVisitId As Long, ByVal strRegNo As String, _
        ByVal lngRegId As Long, ByVal rsAdvice As ADODB.Recordset, ByVal colDiag As Collection) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : GetMainInfo
      ' Author    : YWJ
      ' Date      : 2019-09-23 14:25:53
      ' Parameter :
      '             lngPatiID-����ID
      '             lngVisitId-��ҳID
      '             strRegNo-�Һŵ�
      '             lngRegId -�Һ�ID
      '             colDiag -�������
      ' Purpose   : ��ȡ������Ϣ
      '---------------------------------------------------------------------------------------
      '
          Dim strInfo As String
          
          Dim strDoctor As String
          Dim strDocPost As String   'ҽ��ְ��
          Dim strDoctorList As String
          
          Dim strItemIds As String   '��¼������ĿID
          Dim strIndex As String
          Dim strCondition As String
          
          Dim strKey As String
          Dim strName As String
          Dim strType As String
          Dim strLisIds As String
          Dim strָ��ID As String
          
          Dim i As Long
          Dim lngGroupID As Long
          Dim lngId As Long
          
          Dim blnNext As Boolean
          
          Dim colList As Collection
          Dim colItem As Collection
          Dim colTemp As Collection
          Dim colOther As Collection
          Dim colOtherItem As Collection
          
          Dim arrTemp As Variant
          
          Dim rsDoctor As ADODB.Recordset
          Dim rsItem As ADODB.Recordset
          Dim rsDiag As ADODB.Recordset
          

1         On Error GoTo ErrH
2         Set colList = New Collection
          '��ȡ�����Ϣ
          '����༭���水����¼��Ϊ׼
3         If Not colDiag Is Nothing Then
4             For Each colTemp In colDiag
5                 colList.Add colTemp, "K" & colList.Count
6             Next
7         End If
      '    Set rsDiag = Get������ϼ�¼(lngPatiID, IIf(strRegNo <> "", lngRegId, lngVisitId), IIf(strRegNo <> "", "1,11", "2,12"))
      '    Do While Not rsDiag.EOF
      '        Set colItem = New Collection
      '        colItem.Add NVL(rsDiag!����ID, rsDiag!���ID) & "", "key"
      '        colItem.Add rsDiag!���� & "", "name"
      '        colItem.Add "�����Ϣ", "type"
      '        colList.Add colItem, "K" & colList.Count
      '        rsDiag.MoveNext
      '    Loop
8         With rsAdvice
              '��ҩ��˻���ҩ�о�
9             .Filter = ""
10            For i = 1 To .RecordCount
                  '��ȡ����ҽ��
11                If NVL(!����ҽ��) <> "" Then
12                    strDoctor = NVL(!����ҽ��)
13                    If InStr(strDoctor, "/") > 0 Then strDoctor = Mid(strDoctor, 1, InStr(strDoctor, "/") - 1)
14                    If InStr("," & strDoctorList & ",", "," & strDoctor & ",") = 0 And strDoctor <> "" Then
15                        strDoctorList = strDoctorList & "," & strDoctor
16                    End If
17                End If
                   
18                If InStr("," & strItemIds & ",", "," & !������ĿID & ",") = 0 Then
19                    strItemIds = strItemIds & "," & !������ĿID
20                End If
                   
21                .MoveNext
22            Next
              'ȡרҵ����ְ��
23            If strDoctorList <> "" Then
24                strDoctorList = Mid(strDoctorList, 2)
25                Set rsDoctor = GetRS("��Ա��", "���,����,רҵ����ְ��", strDoctorList, "����", , 1)
26            End If
              '��ȡ������ĿĿ¼
27            If strItemIds <> "" Then
28                strItemIds = Mid(strItemIds, 2)
29                Set rsItem = GetRS("������ĿĿ¼", "ID,����,����", strItemIds)
30            End If
              
31            .Filter = ""
32            strDoctor = ""
33            lngGroupID = 0

34            Do While Not .EOF
35                blnNext = True
36                If strDoctor <> !����ҽ�� & "" Then strDocPost = GetDoctorPost(rsDoctor, NVL(!����ҽ��)): strDoctor = !����ҽ�� & ""
                  'ҩƷ��Ŀ
                  '������Ŀ ��ҩ;��;��ҩƵ��;��������;����;����ҽ��ְ��
37                If InStr(",5,6,7,", "," & !������� & ",") > 0 Then
38                    lngGroupID = Decode(!���ID, 0, !id, !���ID)
39                    strIndex = ""
                      'ѭ������
40                    Do While Not .EOF
41                        If lngGroupID <> Decode(!���ID, 0, !id, !���ID) Then Exit Do
42                        Set colItem = New Collection
43                        If InStr(",5,6,7,", "," & !������� & ",") > 0 Then
44                            colItem.Add !������ĿID & "", "key"
45                            colItem.Add !�걾��λ & "", "name"
46                            colItem.Add "ҩƷ��Ŀ", "type"
                              '������Ŀ
47                            Set colOther = New Collection
48                            Set colOtherItem = New Collection
49                            colOtherItem.Add !ִ��Ƶ�� & "", "��ҩƵ��"
50                            colOtherItem.Add FormatEx(NVL(!��������), 5), "��������"
51                            colOtherItem.Add FormatEx(NVL(!�ܸ�����), 5), "����"
52                            colOtherItem.Add strDocPost, "����ҽ��ְ��"
53                            colOtherItem.Add "��ҩ;��,��ҩƵ��,��������,����,����ҽ��ְ��", "keys"
                              
54                            colOther.Add colOtherItem
55                            colItem.Add colOther, "other"
                              
56                            If strIndex <> "" Then strIndex = strIndex & ","
57                            strIndex = strIndex & "K" & colList.Count
                          
58                        ElseIf !������� & "" = "E" And !id = lngGroupID Then
                              '���Ӹ�ҩ;��
59                            arrTemp = Split(strIndex, ",")
60                            For i = LBound(arrTemp) To UBound(arrTemp)
61                                Set colTemp = colList(arrTemp(i))
62                                For Each colOtherItem In colTemp("other")
63                                    colOtherItem.Add !������ĿID & "", "��ҩ;��"
64                                Next
65                            Next
                              
                              '��ҩ;��
                              '������Ŀ ����ҽ��ְ��
66                            colItem.Add !������ĿID & "", "key"
67                            colItem.Add !ҽ������ & "", "name"
68                            colItem.Add "��ҩ;��", "type"
                              
69                            Set colOther = New Collection
70                            Set colOtherItem = New Collection
                              
71                            colOtherItem.Add strDocPost, "����ҽ��ְ��"
72                            colOtherItem.Add "����ҽ��ְ��", "keys"
                              
73                            colOther.Add colOtherItem
74                            colItem.Add colOther, "other"
75                        End If
76                        colList.Add colItem, "K" & colList.Count
77                        .MoveNext
78                    Loop
79                    blnNext = False 'һ��ҽ���Ѿ������������Ѿ�����¼��ĩβ,��ֹ��������MoveNext
80                ElseIf !������� & "" = "F" Then
                  '������Ŀ
                  '������Ŀ ����ʽ;����ҽ��ְ��
                  '������Ŀ
                  '������Ŀ ����ҽ��ְ��
81                    lngGroupID = Decode(!���ID, 0, !id, !���ID)
82                    strIndex = "": strName = ""
                      'ѭ������
83                    Do While Not .EOF
84                        If lngGroupID <> Decode(!���ID, 0, !id, !���ID) Then Exit Do
85                        lngId = Val(rsAdvice!���ID & "")
86                        Set colItem = New Collection
                                              
87                        If !������� & "" = "F" Then
88                            If lngId = 0 Then
                                  '������
89                                strName = GetItemInfo(rsItem, CLng(!������ĿID & ""))
90                            Else
91                                strName = !ҽ������ & ""
92                            End If
                              
93                            colItem.Add !������ĿID & "", "key"
94                            colItem.Add strName, "name"
95                            colItem.Add "������Ŀ", "type"
              
                              '������Ŀ
96                            Set colOther = New Collection
97                            Set colOtherItem = New Collection
                              
98                            colOtherItem.Add strDocPost, "����ҽ��ְ��"
99                            colOtherItem.Add IIf(lngId = 0, 1, 0), "������" '1-��������;0-��������
100                           colOtherItem.Add "����ʽ,����ҽ��ְ��,������", "keys"
                              
101                           colOther.Add colOtherItem
102                           colItem.Add colOther, "other"
                              
103                           If strIndex <> "" Then strIndex = strIndex & ","
104                           strIndex = strIndex & "K" & colList.Count
                              
105                       ElseIf !������� & "" = "G" Then
                              '��������ʽ
106                           arrTemp = Split(strIndex, ",")
107                           For i = LBound(arrTemp) To UBound(arrTemp)
108                               Set colTemp = colList(arrTemp(i))
109                               For Each colOtherItem In colTemp("other")
110                                   colOtherItem.Add !������ĿID & "", "����ʽ"
111                               Next
112                           Next
                              '������Ŀ
113                           colItem.Add !������ĿID & "", "key"
114                           colItem.Add !ҽ������ & "", "name"
115                           colItem.Add "������Ŀ", "type"
                              
                              '������Ŀ ����ҽ��ְ��
116                           Set colOther = New Collection
117                           Set colOtherItem = New Collection
                              
118                           colOtherItem.Add strDocPost, "����ҽ��ְ��"
119                           colOtherItem.Add "����ҽ��ְ��", "keys"
                              
120                           colOther.Add colOtherItem
121                           colItem.Add colOther, "other"
122                       End If
                          
123                       colList.Add colItem, "K" & colList.Count
                          
124                       .MoveNext
125                   Loop
                      
126                   blnNext = False 'һ��ҽ���Ѿ������������Ѿ�����¼��ĩβ,��ֹ��������MoveNext
                      
127               ElseIf !������� & "" = "K" Then
                      '��Ѫ��Ŀ
                      '������Ŀ ��Ѫ����;��Ѫ��;����ҽ��ְ��
128                   lngGroupID = Decode(!���ID, 0, !id, !���ID)
129                   strKey = "": strName = ""
                      
130                   strIndex = ""
                      'ѭ���������
131                   Do While Not .EOF
132                       If lngGroupID <> Decode(!���ID, 0, !id, !���ID) Then Exit Do
133                       Set colItem = New Collection
134                       lngId = Val(rsAdvice!���ID & "")
135                       If strDoctor <> NVL(!����ҽ��) Then strDocPost = GetDoctorPost(rsDoctor, NVL(!����ҽ��))
136                       If !������� & "" = "K" Then
137                           strName = GetItemInfo(rsItem, CLng(!������ĿID & ""))
138                           colItem.Add !������ĿID & "", "key"
139                           colItem.Add strName, "name"
140                           colItem.Add "��Ѫ��Ŀ", "type"
                              '������Ŀ
141                           Set colOther = New Collection
142                           Set colOtherItem = New Collection
143                           colOtherItem.Add FormatEx(NVL(!�ܸ�����), 5), "��Ѫ��"
144                           colOtherItem.Add strDocPost, "����ҽ��ְ��"
145                           colOtherItem.Add "��Ѫ����,��Ѫ��,����ҽ��ְ��", "keys"
                              
146                           colOther.Add colOtherItem
147                           colItem.Add colOther, "other"
                              
148                           If strIndex <> "" Then strIndex = strIndex & ","
149                           strIndex = strIndex & "K" & colList.Count
                          
150                       ElseIf !������� & "" = "E" Then
                              '������Ѫ����
151                           arrTemp = Split(strIndex, ",")
152                           For i = LBound(arrTemp) To UBound(arrTemp)
153                               Set colTemp = colList(arrTemp(i))
154                               For Each colOtherItem In colTemp("other")
155                                   colOtherItem.Add !������ĿID & "", "��Ѫ����"
156                               Next
157                           Next
                              
                              '��Ѫ����
                              '������Ŀ ����ҽ��ְ��
158                           colItem.Add !������ĿID & "", "key"
159                           colItem.Add !ҽ������ & "", "name"
160                           colItem.Add "��Ѫ����", "type"
                              
161                           Set colOther = New Collection
162                           Set colOtherItem = New Collection
                              
163                           colOtherItem.Add strDocPost, "����ҽ��ְ��"
164                           colOtherItem.Add "����ҽ��ְ��", "keys"
                              
165                           colOther.Add colOtherItem
166                           colItem.Add colOther, "other"

167                       End If
168                       colList.Add colItem, "K" & colList.Count
                          
169                       .MoveNext
170                   Loop
                      
171                   blnNext = False 'һ��ҽ���Ѿ������������Ѿ�����¼��ĩβ,��ֹ��������MoveNext
                      
172               ElseIf !������� & "" = "C" Then
                      '������Ŀ
                      '������Ŀ �ɼ�����;�걾����;����ҽ��ְ��
173                   lngGroupID = Decode(!���ID, 0, !id, !���ID)
174                   strIndex = ""
                      'ѭ������
175                   Do While Not .EOF
176                       If lngGroupID <> Decode(!���ID, 0, !id, !���ID) Then Exit Do
177                       Set colItem = New Collection
178                       If !������� & "" = "C" Then
179                           If strLisIds <> "" Then strLisIds = strLisIds & ","
180                           strLisIds = strLisIds & !������ĿID
                              
181                           colItem.Add !������ĿID & "", "key"
182                           colItem.Add !ҽ������ & "", "name"
183                           colItem.Add "������Ŀ", "type"
                              '������Ŀ
184                           Set colOther = New Collection
185                           Set colOtherItem = New Collection
186                           colOtherItem.Add !�걾��λ & "", "�걾����"
187                           colOtherItem.Add strDocPost, "����ҽ��ְ��"
188                           colOtherItem.Add "�ɼ�����,�걾����,����ҽ��ְ��", "keys"
                              
189                           colOther.Add colOtherItem
190                           colItem.Add colOther, "other"
                              
191                           If strIndex <> "" Then strIndex = strIndex & ","
192                           strIndex = strIndex & "K" & colList.Count
                          
193                       ElseIf !������� & "" = "E" And !id = lngGroupID Then
                              '���Ӳɼ�����
194                           arrTemp = Split(strIndex, ",")
195                           For i = LBound(arrTemp) To UBound(arrTemp)
196                               Set colTemp = colList(arrTemp(i))
197                               For Each colOtherItem In colTemp("other")
198                                   colOtherItem.Add !������ĿID & "", "�ɼ�����"
199                               Next
200                           Next
                              
                              '�ɼ�����
                              '������Ŀ ����ҽ��ְ��
201                           colItem.Add !������ĿID & "", "key"
202                           colItem.Add GetItemInfo(rsItem, CLng(!������ĿID & "")), "name"
203                           colItem.Add "�ɼ�����", "type"
                              
204                           Set colOther = New Collection
205                           Set colOtherItem = New Collection
                              
206                           colOtherItem.Add strDocPost, "����ҽ��ְ��"
207                           colOtherItem.Add "����ҽ��ְ��", "keys"
                              
208                           colOther.Add colOtherItem
209                           colItem.Add colOther, "other"
210                       End If
211                       colList.Add colItem, "K" & colList.Count
212                       .MoveNext
213                   Loop
214                   blnNext = False 'һ��ҽ���Ѿ������������Ѿ�����¼��ĩβ,��ֹ��������MoveNext
215               ElseIf !������� & "" = "D" Then
                  '�����Ŀ
                  '������Ŀ ��λ;����;����ҽ��ְ��
216                   lngGroupID = Decode(!���ID, 0, !id, !���ID)
217                   strName = ""
218                   Set colItem = New Collection
219                   Set colOther = New Collection
                      'ѭ���������
220                   Do While Not .EOF
221                       If lngGroupID <> Decode(!���ID, 0, !id, !���ID) Then Exit Do
222                       lngId = Val(rsAdvice!���ID & "")
223                       If lngId = 0 Then
224                           colItem.Add !������ĿID & "", "key"
225                           colItem.Add "�����Ŀ", "type"
226                           strName = !ҽ������ & ""
227                       Else
228                           strName = !ҽ������ & ""
                              '������Ŀ
229                           Set colOtherItem = New Collection
230                           colOtherItem.Add !�걾��λ & "", "��λ"
231                           colOtherItem.Add !��鷽�� & "", "����"
232                           colOtherItem.Add strDocPost, "����ҽ��ְ��"
233                           colOtherItem.Add "��λ,����,����ҽ��ְ��", "keys"
                              
234                           colOther.Add colOtherItem
235                       End If
236                       .MoveNext
237                   Loop
                      
238                   blnNext = False 'һ��ҽ���Ѿ������������Ѿ�����¼��ĩβ,��ֹ��������MoveNext
239                   colItem.Add strName, "name"
                      
240                   If colOther.Count = 0 Then '�������ҽ�����⴦��
241                       Set colOtherItem = New Collection
242                       colOtherItem.Add strDocPost, "����ҽ��ְ��"
243                       colOtherItem.Add "����ҽ��ְ��", "keys"
244                   End If
245                   colItem.Add colOther, "other"
                      
246                   colList.Add colItem, "K" & colList.Count
247               Else
                      '����ָ��
                      '������Ŀ ָ������
                      
                      '����ҽ����Ŀ
                      '������Ŀ ����ҽ��ְ��
248                   If !������� & "" = "H" Then
                          '������Ŀ
                          '������Ŀ ����ҽ��ְ��
249                       strType = "������Ŀ"
250                   ElseIf !������� & "" = "E" Then
                          '������Ŀ
                          '������Ŀ ����ҽ��ְ��
251                       strType = "������Ŀ"
252                   Else
253                       strType = "����ҽ����Ŀ"
254                   End If
255                   Set colItem = New Collection
256                   Set colOther = New Collection
                      
257                   colItem.Add !������ĿID & "", "key"
258                   colItem.Add !ҽ������ & "", "name"
259                   colItem.Add strType, "type"
                          
260                   Set colOtherItem = New Collection
261                   colOtherItem.Add strDocPost, "����ҽ��ְ��"
262                   colOtherItem.Add "����ҽ��ְ��", "keys"
                      
263                   colOther.Add colOtherItem
                      
264                   colItem.Add colOther, "other"
265                   colList.Add colItem, "K" & colList.Count
266               End If
                                      
267               If blnNext Then .MoveNext
268           Loop
              'ָ��ID,����ָ��,skey,sname
269           If strLisIds <> "" And Not gobjLIS Is Nothing Then
270               Set rsItem = Nothing
271               On Error Resume Next
272               Set rsItem = gobjLIS.GetGroupItemInfo(strLisIds)
273               On Error GoTo ErrH '������һ���Դ���,���¿������󲶻�
274               If Not rsItem Is Nothing Then
275                   strָ��ID = ""
276                   Do While Not rsItem.EOF
277                       If strָ��ID <> rsItem!ָ��ID & "" Then
278                           If strָ��ID <> "" Then
279                               colItem.Add colOther, "other"
280                               colList.Add colItem, "K" & colList.Count
                                  
281                               Set colItem = Nothing
282                               Set colOther = Nothing
283                           End If
284                           strָ��ID = rsItem!ָ��ID & ""
                              
285                           Set colItem = New Collection
286                           colItem.Add rsItem!ָ��ID & "", "key"
287                           colItem.Add rsItem!����ָ�� & "", "name"
288                           colItem.Add "����ָ��", "type"
                               
289                           If rsItem!sname & "" <> "" And rsItem!sKey & "" <> "" Then
290                               Set colOther = New Collection
291                               Set colOtherItem = New Collection
                                  
292                               colOtherItem.Add rsItem!sname & "", rsItem!sKey & ""
293                               colOtherItem.Add rsItem!sKey & "", "keys"
                                  
294                               colOther.Add colOtherItem
295                           End If
296                       Else
297                           If rsItem!sname & "" <> "" And rsItem!sKey & "" <> "" And Not colOther Is Nothing Then
298                               Set colOtherItem = New Collection
                                  
299                               colOtherItem.Add rsItem!sname & "", rsItem!sKey & ""
300                               colOtherItem.Add rsItem!sKey & "", "keys"
                                  
301                               colOther.Add colOtherItem
302                           End If
303                       End If
                      
304                       rsItem.MoveNext
305                       If rsItem.EOF Then
306                           colItem.Add colOther, "other"
307                           colList.Add colItem, "K" & colList.Count
308                       End If
309                   Loop
310               End If
311           End If
                      
312       End With
313       If colList.Count > 0 Then
314           strInfo = GetMainJson(colList)
315       End If
316       GetMainInfo = strInfo

317       Exit Function

ErrH:
318       MsgBox "��zlCISRule.mdlPublic.GetMainInfo�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetDoctorPost(ByVal rsDoctor As ADODB.Recordset, ByVal strDoctor As String) As String
'����:��ȡҽʦרҵ����ְ��
'����:
    rsDoctor.Filter = "����='" & strDoctor & "'"
    If Not rsDoctor.EOF Then GetDoctorPost = NVL(rsDoctor!רҵ����ְ��)
End Function

Private Function GetItemInfo(ByVal rsItem As ADODB.Recordset, ByVal lngId As Long) As String
'����:��ȡҽʦרҵ����ְ��
'����:
    rsItem.Filter = "ID=" & lngId & ""
    If Not rsItem.EOF Then GetItemInfo = NVL(rsItem!����)
End Function

Public Function GetMainJson(ByVal colList As Collection) As String
      '���ܣ�����������Ϣ
          Dim colItem As Collection
          Dim colOther As Collection
          Dim colOtherItem As Collection
          Dim arrKeys As Variant
          
          Dim strInfo As String
          Dim strTemp As String
          Dim strOther As String
          
          Dim i As Long

1         On Error GoTo ErrH

2         strInfo = ""
3         For Each colItem In colList
4             Set colOther = GetCollValue(colItem, "other")
5             strOther = ""
6             If Not colOther Is Nothing Then
7                 For Each colOtherItem In colOther
8                     arrKeys = Split(GetCollElement(colOtherItem, "keys"), ",")
9                     strTemp = ""
10                    For i = LBound(arrKeys) To UBound(arrKeys)
11                        If strTemp <> "" Then strTemp = strTemp & ","
12                        If arrKeys(i) <> "" Then
13                           strTemp = strTemp & "{\""skey\"":\""" & arrKeys(i) & "\"",\""sname\"":\""" & GetCollElement(colOtherItem, CStr(arrKeys(i))) & "\""}"
14                        End If
15                    Next
16                    If strOther <> "" Then strOther = strOther & ","
17                    strOther = strOther & "{\""value_group\"":[" & strTemp & "]}"
18                Next
19                If strOther <> "" Then strOther = "\""condition_info\"":[" & strOther & "]"
21            End If
22            If strInfo <> "" Then strInfo = strInfo & ","
23            strInfo = strInfo & "{\""key\"":\""" & GetCollElement(colItem, "key") & "\""," & vbNewLine & _
                                  "\""name\"":\""" & GetCollElement(colItem, "name") & "\""," & vbNewLine & _
                                  "\""type\"":\""" & GetCollElement(colItem, "type") & "\"""
24            If strOther = "" Then
25                strInfo = strInfo & "}"
26            Else
27                strInfo = strInfo & "," & strOther & "}"
28            End If
                                   
29        Next
30        strInfo = "\""main_info\"":[" & strInfo & "]"
31        GetMainJson = strInfo
32        Exit Function

ErrH:
33        MsgBox "��zlCISRule.mdlPublic.GetMainJson�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Function TestJson(ByVal bytFunc As Byte) As String
    Dim strJson As String
    If bytFunc = 1 Then
        strJson = "{\""cdss_in\"":{\""patient_info\"":{\""pid\"": \""5066404\"",\""visit_id\"": \""1\""," & _
                "\""visit_no\"": \""314929\"",\""name\"": \""������\"",\""age\"": \""31��\"",\""gender\"": \""Ů\"",\""marital_status\"": \""�ѻ�\"",\""operator_id\"": \""489b7bba-31cd-4f59-8fef-c12f0570db61\"",\""operator\"": \""֣־��\"",\""enc_type\"": \""2\"",\""scene\"": \""1\""}," & _
                "\""main_info\"": [{\""key\"": \""168\"",\""name\"": \""ע����������ù��\"",\""type\"": \""ҩƷ��Ŀ\"",\""condition_info\"":[{\""value_group\"":[{\""skey\"": \""��ҩ;��\"",\""sname\"": \""2203\""},{\""skey\"": \""��ҩƵ��\"",\""sname\"": \""1\""},{\""skey\"": \""��������\"",\""sname\"": \""1\""},{\""skey\"": \""����\"",\""sname\"": \""1\""},{\""skey\"": \""������\"",\""sname\"": \""֣־��\""}]}]},{\""key\"": \""2203\"",\""name\"": \""����ע��\"",\""type\"": \""��ҩ;��\""}]}}"
        TestJson = "{""�ӿ�json_in"":""" & strJson & """}"
        
    ElseIf bytFunc = 2 Then
        strJson = "{" & vbNewLine & _
                """businss"":""F8E7C2918A6C4060B29FE5D3FD66135A""," & vbNewLine & _
                """inquiry"":[{" & vbNewLine & _
                "    ""observ_item_id"":""F0E3D17C3CDE4FBF89FF020372D0A1EF""," & vbNewLine & _
                "    ""item_name"":""��������""," & vbNewLine & _
                "    ""item_code"":""""," & vbNewLine & _
                "    ""observ_item_values"":[{" & vbNewLine & _
                "        ""item_detail_id"":""A43FD7C4166A470E9CD266FC9AC9D0B3""," & vbNewLine & _
                "        ""disp_name"":""��""," & vbNewLine & _
                "        ""default_sign"":""1""" & vbNewLine & _
                "        }, {" & vbNewLine & _
                "        ""item_detail_id"":""56E11442BA374FA8B7DF2FECE895AA13""," & vbNewLine & _
                "        ""disp_name"":""��""," & vbNewLine & _
                "        ""default_sign"":""0""" & vbNewLine & _
                "        }]" & vbNewLine & _
                "    }, {"
        strJson = strJson & "" & vbNewLine & _
                """observ_item_id"":""F14CEA53C2B646A399FD0DD491BAF0FE""," & vbNewLine & _
                """item_name"":""��������""," & vbNewLine & _
                """item_code"":""""," & vbNewLine & _
                """observ_item_values"":[{" & vbNewLine & _
                "    ""item_detail_id"":""77F1DF1C24884576BFBCCBDA39DA3D9F""," & vbNewLine & _
                "    ""disp_name"":""��������""," & vbNewLine & _
                "    ""default_sign"":""0""" & vbNewLine & _
                "    }, {" & vbNewLine & _
                "    ""item_detail_id"":""14345A4091194C4FAD518BDB029B1325""," & vbNewLine & _
                "    ""disp_name"":""��������""," & vbNewLine & _
                "    ""default_sign"":""0""" & vbNewLine & _
                "    }, {" & vbNewLine & _
                "    ""item_detail_id"":""8A64280BC22D432CB28346EA653E58F8""," & vbNewLine & _
                "    ""disp_name"":""��������""," & vbNewLine & _
                "    ""default_sign"":""0""" & vbNewLine & _
                "    }]" & vbNewLine & _
                "}]," & vbNewLine & _
                ""
        strJson = strJson & "" & vbNewLine & _
                """messages"":[{" & vbNewLine & _
                "    ""business_name"":""ҩƷ����""," & vbNewLine & _
                "    ""return_info"":""���ˡ������ա�ʹ��ҩƷ��������ù��ע���������ֹ������ע�䡿""," & vbNewLine & _
                "    ""key"":""2""," & vbNewLine & _
                "    ""name"":""������ù��ע���""," & vbNewLine & _
                "    ""rule_name"":""��ҩ����""," & vbNewLine & _
                "    ""taboo_level"":""����""" & vbNewLine & _
                "    }]" & vbNewLine & _
                "}"
        TestJson = """out"":" & strJson
    ElseIf bytFunc = 3 Then
        '��Ԥ����
        strJson = "{""businss"":""DA32302FD74541B98DC4F3F992E5B206"",""messages"":[{""business_name"":""ҩƷ����"",""return_info"":""���ˡ������ա�����š�314929�����򡾲����ڽ�ֹʹ�á���ע����������ù�ء�"",""key"":""168"",""name"":""ע����������ù��"",""rule_name"":""�����ڽ�ֹʹ��"",""taboo_level"":""��ֹ"",""class_id"":""46dfdfbd-fdab-429a-877c-a4dee02752e7"",""detail_id"":""b94b6533-d96e-4836-a386-825273461144""}]}"
        TestJson = strJson
    ElseIf bytFunc = 4 Then
        '��������
        strJson = "{""cdss_in"":{""patient_info"":{""pid"": ""4613704"",""visit_id"": ""1"",""visit_no"": ""303740"",""name"": ""������"",""age"": ""71��"",""birthday"": ""1948-03-12"",""gender"": ""Ů"",""marital_status"": ""�ѻ�"",""operator_id"": ""489b7bba-31cd-4f59-8fef-c12f0570db61"",""operator"": ""֣־��"",""enc_type"": ""2""},""main_info"": [{""key"": ""168"",""name"": ""ע����������ù��"",""type"": ""ҩƷ��Ŀ"",""condition_info"":[{""value_group"":[{""skey"": ""��ҩ;��"",""sname"": ""144511""},{""skey"": ""��ҩƵ��"",""sname"": ""1""},{""skey"": ""��������"",""sname"": ""1""},{""skey"": ""����"",""sname"": ""1""},{""skey"": ""������"",""sname"": ""֣־��""}]}]},{""key"": ""144511"",""name"": ""����ע��"",""type"": ""��ҩ;��""}]}}"
        strJson = Replace(strJson, """", "\""")
        TestJson = "{""�ӿ�json_in"":""" & strJson & """}"
    ElseIf bytFunc = 5 Then
     '�����������
        strJson = "{""businss"":""97CEF2E3E5894591A4EB1BA3A215146E"",""inquiry"":" & vbNewLine & _
                "[{""observ_item_id"":""F0E3D17C3CDE4FBF89FF020372D0A1EF"",""item_name"":""��������"",""item_code"":"""",""observ_item_values"":" & vbNewLine & _
                "[{""item_detail_id"":""A43FD7C4166A470E9CD266FC9AC9D0B3"",""disp_name"":""��"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""56E11442BA374FA8B7DF2FECE895AA13"",""disp_name"":""��"",""default_sign"":""0""}]}," & vbNewLine & _
                "{""observ_item_id"":""9EF6C13B62094E698E481868C703E634"",""item_name"":""����״̬"",""item_code"":"""",""observ_item_values"":" & vbNewLine & _
                "[{""item_detail_id"":""427EAB8923FF49258FC935E74BAACAE7"",""disp_name"":""��"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""CA018CB5A6914896AA4FC4D1C137164F"",""disp_name"":""��"",""default_sign"":""0""}]},"
        
        strJson = strJson & "{""observ_item_id"":""4BA5C77FE9D9408389A9BE085E30E99A"",""item_name"":""�ι��ܲ�ȫ"",""item_code"":""""," & vbNewLine & _
                """observ_item_values"":[{""item_detail_id"":""C532D50ACF6C4DDC8D20A0039B87C2FD"",""disp_name"":""��"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""1D958F80177546F0BC06619AEDCF00A6"",""disp_name"":""��"",""default_sign"":""0""}]}," & vbNewLine & _
                "{""observ_item_id"":""1640B9E956AD4372BC353DE709455B45"",""item_name"":""�����ܲ�ȫ"",""" & vbNewLine & _
                "item_code"":"""",""observ_item_values"":[{""item_detail_id"":""9EE1C7FC4EB84B28A829CC1B9654A4F1"",""disp_name"":""��"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""562F75B5A00849B5A5A92C0144CD49BC"",""disp_name"":""��"",""default_sign"":""0""}]}],"
        
        strJson = strJson & """messages"":[{""business_name"":""ҩƷ����"",""return_info"":""���ˡ�������������š�303740�����򡾲����ڽ�ֹʹ�á���ע����������ù�ء�""," & vbNewLine & _
                """key"":""168"",""name"":""ע����������ù��"",""rule_name"":""�����ڽ�ֹʹ��"",""taboo_level"":""��ֹ""," & vbNewLine & _
                """class_id"":""46dfdfbd-fdab-429a-877c-a4dee02752e7"",""detail_id"":""b94b6533-d96e-4836-a386-825273461144""}," & vbNewLine & _
                "{""business_name"":""ҩƷ����"",""return_info"":""���ˡ�������������š�303740��ʹ��ҩƷ��ע����������ù�ء������á�����ע�䡿""," & vbNewLine & _
                """key"":""168"",""name"":""ע����������ù��"",""rule_name"":""��ҩ����"",""taboo_level"":""����""," & vbNewLine & _
                """class_id"":""46dfdfbd-fdab-429a-877c-a4dee02752e7"",""detail_id"":""72be25bb-0a1e-4aed-ab0f-ce181d87a694""}]}"
                
        TestJson = strJson
    End If
End Function

Public Function GetCollValue(ByVal colList As Collection, ByVal varRow As Variant, Optional ByVal strElement As String) As Variant
'���ܣ���ȡJson���鷵�صļ���������ָ���л�ָ��Ԫ�ص�ֵ
'������
'  varRow=���������йؼ���
'  strElement=Ԫ����
'���أ�
'  ��δ����strElement����ʱ������ָ���еļ��϶��󣻵�����strElement����ʱ������ָ����ָ��Ԫ�ص�ֵ
'  ʧ��ʱ����Nothing��Empty�������ᱨ��

    If strElement <> "" Then
        GetCollValue = Empty
    Else
        Set GetCollValue = Nothing
    End If
    
    If colList Is Nothing Then Exit Function
    
    On Error Resume Next
    If strElement <> "" Then
        GetCollValue = colList(varRow)(strElement)
    Else
        Set GetCollValue = colList(varRow)
    End If
    Err.Clear: On Error GoTo 0
End Function

Public Function GetCollElement(ByVal colList As Collection, ByVal strElement As String) As Variant
'���ܣ���ȡ���������е�Ԫ��ֵ(Ԫ��ֵΪ������������)
'������
'  varRow=���������йؼ���
'  strElement=Ԫ����
'���أ�
'   ����ָ����ָ��Ԫ�ص�ֵ
'   ʧ��ʱ����Empty

 
    GetCollElement = Empty
    If colList Is Nothing Then Exit Function
    On Error Resume Next
    GetCollElement = colList(strElement)
    Err.Clear: On Error GoTo 0
End Function

Public Function HandleMessage(ByVal colList As Collection) As Boolean
'����:���ݷ��ؾ�ʾ����Ե�ǰ�������и�Ԥ��
'����ֵ:T-��ֹ��ǰ����;F-������ǰ����
    Dim i As Long
    Dim strMsg As String
    Dim lngLevel As Byte '1-����;2-����;3-��ֹ
    Dim lngMaxLevel As Byte
    Dim blnRet As Boolean
    
    For i = 1 To colList.Count
        If strMsg <> "" Then strMsg = strMsg & vbCrLf
        strMsg = strMsg & colList(i)("return_info")
        lngLevel = Decode(CStr(colList(i)("taboo_level")), "��ֹ", 3, "����", 2, "����", 1, 0)
        If lngLevel > lngMaxLevel Then lngMaxLevel = lngLevel
    Next
    Select Case lngMaxLevel
    
    Case 1
        MsgBox strMsg, vbInformation + vbOKOnly, gstrSysName
    Case 2
        If MsgBox(strMsg & vbCrLf & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            blnRet = True
        End If
    Case 3
        MsgBox strMsg, vbExclamation + vbOKOnly, gstrSysName
        blnRet = True
    End Select
    HandleMessage = blnRet
End Function

Public Function GetRS(ByVal strTableName As String, ByVal strFileds As String, ByVal strInput As String, _
        Optional ByVal strWhere As String = "ID", Optional ByVal bytModel As Byte = 0, Optional ByVal bytType As Byte = 0) As Variant
'����:����ָ����ָ���ֶεļ�¼��
'������strTableName-����
'     strFileds
'     strInput ��ʽ1(1����������)��ID1,ID2,...
'              ��ʽ2(2����������)������1,��Χ1;����2,��Χ2;...
'             strSQL = "Select ����, ����, ���÷�Χ" & vbNewLine & _
'                "From ����Ƶ����Ŀ" & vbNewLine & _
'                "Where (����, ���÷�Χ) In (Select /*+cardinality(B,10)*/" & vbNewLine & _
'                "                      C1, C2" & vbNewLine & _
'                "                     From Table(f_Str2list2('ÿ�����,1|ÿ������,1', ';', ',')) B)"
'    bytModel=1 ��������Ϊ����
'    ��bytModel=1ʱ�� bytType=0-����� C1,C2 ͬΪ�ַ��� =1-C1(Number),C2(Number);=2-C1(char),C2(Number);=3-C1(Number),C2(Char)
'    ��bytModel=0ʱ�� bytType=0-f_num2list; bytType=1 f_Str2list


    Dim strSQL As String
    Dim strSub As String
    Dim strFun As String
    Dim arrTmp As Variant
    
    On Error GoTo ErrH
    
    If bytModel = 1 Then
        If bytType = 0 Then
            strSub = " C1,C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 1 Then
            strSub = " C1,C2 "
            strFun = "f_num2list2"
        ElseIf bytType = 2 Then
            strSub = "C1,To_Number(C2) As C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 3 Then
            strSub = " To_Number(C1) As C1,C2 "
            strFun = "f_Str2list2"
        End If
        strSQL = " Select  " & strFileds & vbNewLine & _
                " From  " & strTableName & vbNewLine & _
                " Where (" & strWhere & ") In (Select /*+cardinality(B,10)*/" & vbNewLine & _
                "                    " & strSub & vbNewLine & _
                "                     From Table(" & strFun & "([1], ';', ',')) B)"
    Else
        If bytType = 0 Then
            strFun = "f_num2list"
        ElseIf bytType = 1 Then
            strFun = "f_Str2list"
        End If
        arrTmp = Split(strInput, ",")
        If UBound(arrTmp) = 0 Or strInput = "" Then
            strSQL = "Select " & strFileds & "  From " & strTableName & " Where " & strWhere & " = [1]"
        ElseIf UBound(arrTmp) > 0 Then
            strSQL = "Select " & strFileds & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where " & strWhere & " In (Select /*+cardinality(A,10)*/ * From Table(" & strFun & "([1]))A )"
        End If
    End If
    Set GetRS = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", strInput)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������ϼ�¼(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str���� As String) As ADODB.Recordset
'���ܣ���ȡ������ϼ�¼
'������lng����ID�����ﲡ�˴��Һ�ID��סԺ���˴���ҳID
'       �������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'       ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);
    Dim strSQL As String

    On Error GoTo ErrH
    strSQL = "Select a.ID,a.����id, a.���id, a.�������, a.��ϴ���, Nvl(b.����, c.����) As ����, NVL(Nvl(b.����, c.����),a.�������) ����" & vbNewLine & _
             ",a.��¼����,a.��¼�� " & vbNewLine & _
             "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & vbNewLine & _
             "Where a.����id = [1] And a.��ҳid = [2] And ȡ��ʱ�� Is Null And ��¼��Դ IN (1, 3) And Instr(',' ||[3]|| ',', ',' || ������� || ',') > 0 And a.����id = b.Id(+) And" & vbNewLine & _
             "      a.���id = c.Id(+)" & vbNewLine & _
             "Order By ��¼��Դ, �������, ��ϴ���"
    Set Get������ϼ�¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng����ID, str����)

    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zlPublicHisCommLis.clsPublicHisCommLis")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub
