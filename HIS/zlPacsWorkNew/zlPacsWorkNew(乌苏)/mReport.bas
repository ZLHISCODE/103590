Attribute VB_Name = "mReport"
Option Explicit

Public pReport_CheckViewName As String
Public pReport_ResultName As String
Public pReport_AdviceName As String

Public preWinProc As Long
Public fReport As frmReportWord

Public Const ReportViewType_������� = "�������"
Public Const ReportViewType_������ = "������"
Public Const ReportViewType_���� = "����"
Public Const ReportViewType_������� = "�������"
Public Const ReportViewType_��첿λ = "��첿λ"

'################################################################################################################
'## ���ܣ�  �ж�ָ���û��Ƿ�������ҽʦ
'##
'## ������  lngUserID       ���û�ID
'##         strUserName     ���û���
'##         lngPatiID       ������ID
'##         lngPatiPageID   ����ҳID
'##
'## ˵����  ���ݡ���Ա���еġ�Ƹ�μ���ְ���ֶ�ȷ��ҽ����������סԺҽʦ������ҽʦ������ҽʦ��
'##         �����˱䶯��¼�е�ҽ�����𣬴Ӷ�ȷ����˼���
'################################################################################################################
Public Function GetUserSignLevel(lngUserID As Long, Optional strUserName As String, _
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevelEnum
    Dim RS As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    err = 0: On Error GoTo errHand
    gstrSQL = "Select g.����" & vbNewLine & _
            "From zlRoleGrant g, Sys.Dba_Role_Privs r, �ϻ���Ա�� p" & vbNewLine & _
            "Where r.Grantee = p.�û��� And g.��ɫ = r.Granted_Role And g.ϵͳ = [2] And g.��� = [3] And g.���� = [4] And" & vbNewLine & _
            "      p.��Աid = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select [4] As ���� From �ϻ���Ա�� p Where �û��� = '" & UCase(UserInfo.�û���) & "' And p.��Աid = [1]"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "mReport", lngUserID, glngSys, 1070, "ǩ��Ȩ")
    If RS.RecordCount <= 0 Then GetUserSignLevel = cprSL_�հ�: Exit Function
    
    gstrSQL = "select Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not RS.EOF Then
        lngR = Nvl(RS("Ƹ�μ���ְ��"), 0)
    End If
    Select Case lngR    '1 ����  2 ����  3 �м�  4 ����/ʦ��  5 Ա/ʿ  9 ��Ƹ
    Case 1: lngLevel1 = cprSL_����
    Case 2: lngLevel1 = cprSL_����
    Case 3: lngLevel1 = cprSL_����
    Case Else: lngLevel1 = cprSL_����
    End Select
    RS.Close
    
    If lngPatiID > 0 Then
        gstrSQL = "Select ����ҽʦ, ����ҽʦ, ����ҽʦ " & _
            " From ���˱䶯��¼ " & _
            " Where ����ID = [1] And ��ҳID = [2] And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) " & _
            "       And ��ʼʱ�� Is Not Null And Nvl(���Ӵ�λ, 0) = 0"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If RS.EOF Then
            lngLevel2 = cprSL_����
        Else
            If RS.Fields("����ҽʦ") = IIf(strUserName = "", UserInfo.����, strUserName) Then
                lngLevel2 = cprSL_����
            ElseIf RS.Fields("����ҽʦ") = IIf(strUserName = "", UserInfo.����, strUserName) Then
                lngLevel2 = cprSL_����
            Else
                lngLevel2 = cprSL_����
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

errHand:
    GetUserSignLevel = cprSL_�հ�
End Function

'################################################################################################################
'## ���ܣ�  ���������ı�����ָ���ؼ�������Ķ�λ��Ϣ
'##
'## ������  edtThis         :   IN  ���༭�ؼ�
'##         strKeyType      :   IN  �������ؼ������ơ�ȡֵΪ��"O"��"P"��"T"��"E"��"U"
'##         lngKey           :   IN  �����������ҵĹؼ���ID�š�
'##         lngKSS��lngKSE  :   OUT ���ֱ��ʾ��ʼ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         lngKES��lngKEE  :   OUT ���ֱ��ʾ��ֹ�ؼ��ֵĿ�ʼλ�úͽ���λ�ã�
'##         blnNeeded:      :   OUT ���Ƿ��Ǳ�������
'##
'## ���أ�  ����ҵ��ùؼ��־���λ�ã��򷵻�True�����򷵻�False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '��������.Text���ԣ������һ���ַ�������������ʱ�俪֧��
    
    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   'ֻ��ȡ.Text����1�Σ�����
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '���Ƿ��ǹؼ���
            If .TOM.TextDocument.range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                i = i + 1
                GoTo LL1
            End If
            '���ҵ���ʼ�ؼ���
            
            '���ҽ����ؼ���
            j = i + 16
LL2:
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '���Ƿ��ǹؼ���
                If .TOM.TextDocument.range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '�ҵ������ؼ���
                strKeyType = strKeyType
                lngKSS = i - 1 'ת��Ϊ0��ʼ������λ�á�
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function


Public Sub richTextBoxShowElements(rText As RichTextBox)
    Dim strSel As String
    Dim miESingleS As Integer
    Dim miESingleE As Integer
    Dim miEMultiS As Integer
    Dim miEMultiE As Integer
    
    
    '�жϵ�ǰѡ�������Ƿ�Ҫ��
    If rText.SelColor = vbBlue Then
        miESingleS = InStrRev(rText.Text, "{{", rText.SelStart, vbTextCompare)
        miEMultiS = InStrRev(rText.Text, "{<", rText.SelStart, vbTextCompare)
        If miESingleS > miEMultiS Then  '��ǰ��ӽ������ǵ�ѡҪ��
            miESingleE = InStr(rText.SelStart, rText.Text, "}}", vbTextCompare)
            miESingleE = miESingleE + 1
            If miESingleE > miESingleS Then
                '�ǵ�ѡҪ��
                strSel = Left(rText.Text, miESingleE)
                strSel = Right(strSel, miESingleE - miESingleS + 1)
                frmReportElement.ShowElement strSel, 0
                rText.SelStart = miESingleS - 1
                rText.SelLength = miESingleE - miESingleS + 1
                rText.SelText = frmReportElement.strReturnElement
            End If
        ElseIf miEMultiS > miESingleS Then  '��ǰ��ӽ����Ƕ�ѡҪ��
            miEMultiE = InStr(rText.SelStart, rText.Text, ">}", vbTextCompare)
            miEMultiE = miEMultiE + 1
            If miEMultiE > miEMultiS Then
                '�Ƕ�ѡҪ��
                strSel = Left(rText.Text, miEMultiE)
                strSel = Right(strSel, miEMultiE - miEMultiS + 1)
                frmReportElement.ShowElement strSel, 1
                rText.SelStart = miEMultiS - 1
                rText.SelLength = miEMultiE - miEMultiS + 1
                rText.SelText = frmReportElement.strReturnElement
            End If
        Else    '����Ҫ�ص�λ����ȣ�˵��������0����ǰʲôҪ�ض�û��
        
        End If
    End If
End Sub

Public Function Wndproc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer
    On Error Resume Next
    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)
    Select Case Msg
        Case WM_MOUSEWHEEL
            If fReport.picWordShow.Visible = False Or fReport.vscroWordH.Enabled = False Then Exit Function
            
            If Sgn(wzDelta) = 1 Then
                If fReport.vscroWordH.value - 1 < 0 Then
                    fReport.vscroWordH.value = 0
                Else
                    fReport.vscroWordH.value = fReport.vscroWordH.value - 1
                End If
            Else
                If fReport.vscroWordH.value + 1 > fReport.vscroWordH.Max Then
                    fReport.vscroWordH.value = fReport.vscroWordH.Max
                Else
                    fReport.vscroWordH.value = fReport.vscroWordH.value + 1
                End If
            End If
    End Select
    Wndproc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function

Public Function zlGetWordPower() As Integer
'******************************************************************************************************************
'���ܣ���õ�ǰ�û��Ĵʾ�����Ȩ��
'���أ��ʾ����Ȩ����ֵ
'******************************************************************************************************************
    Dim intWordPower As Integer
    Dim strPrivs As String
    
    strPrivs = GetPrivFunc(glngSys, 1070)
    If InStr(1, strPrivs, "ȫԺ�����ʾ�") <> 0 Then
        intWordPower = 0
    ElseIf InStr(1, strPrivs, "���Ҳ����ʾ�") <> 0 Then
        intWordPower = 1
    ElseIf InStr(1, strPrivs, "���˲����ʾ�") <> 0 Then
        intWordPower = 2
    Else
        intWordPower = -1
    End If
    zlGetWordPower = intWordPower
End Function

Public Function zlDefaultWordCode(lngClassID As Long) As String
'���ܣ����ôʾ�ʾ����Ĭ�ϱ��
'������ lngClassID --- �ʾ����ID

    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    strSql = "Select LPad(Nvl(To_Number(Max(���)), 0) + 1, Nvl(Max(Length(���)), 5), '0') As ����" & vbNewLine & _
            "From �����ʾ�ʾ��" & vbNewLine & _
            "Where ����id = [1]"
    err = 0: On Error Resume Next
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ʾ���", lngClassID)
    zlDefaultWordCode = rsTemp.Fields(0).value
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSignSourceString(int��ȡ���� As Integer, lngReportID As Long, intǩ���汾 As Integer, blnMoved As Boolean, _
    thisSign As cEPRSign, strSourceOut As String) As Integer
'------------------------------------------------
'���ܣ���ȡ���ڵ���ǩ����ǩ����֤�ı���Դ������
'������ int��ȡ���� -- 1��ǩ��ʱ��ȡԴ�ģ�2��ǩ����֤ʱ��ȡԴ��
'       lngReportID -- ����ID�����Ӳ�����¼ID
'       intǩ���汾 -- ����ǩ��/��֤ǩ����ȡԴ�ĵİ汾��
'       blnMoved --- ���������Ƿ��Ѿ�ת��
'       thisSign --- ǩ������ǩ����ʱ����˶�����֤ǩ����ʱ����nothing
'       strSourceOut -- �����ء�ǩ��Դ��
'���أ� ǩ��/��֤ǩ����Դ�����ɹ���
'-----------------------------------------------
    Dim intRule As Integer
    Dim lngǩ��ID  As Long                  'ǩ�����ڵ��е�ID
    Dim strSql As String
    Dim rs������¼ As ADODB.Recordset
    Dim rs�������� As ADODB.Recordset
    Dim rsǩ����¼ As ADODB.Recordset
    Dim strǩ��ʱ�� As String
    Dim arr��������() As String
    
    'Դ����ȡ����
    'intRule = 1ʱ����ȡ ID������ID��Ӥ���������ˣ�����ʱ�䣬ҽ��������ǩ������ǩ��ʱ��,���������������������
    '��֤ǩ����ʱ��ҽ��������ǩ������ǩ��ʱ���ǩ����¼�л�ȡ���ֱ���ҽ������= �������ı�����ǩ������=��Ҫ�ر�ʾ����ǩ��ʱ�� =���������ԣ�5����
    'ǩ����ʱ��ҽ��������ǩ������ǩ��ʱ�� ��ǩ�������л�ȡ
    On Error GoTo err
    
    If lngReportID = 0 Or intǩ���汾 = 0 Then Exit Function
    
    
    '��ʼ��Ĭ��ֵ
    intRule = 1
    strSourceOut = ""
    
    '����int��ȡ���� ���ж���ǩ��������֤ǩ�����ֱ�Ӷ�Ӧ�ĵط���ȡ����
    '�ӵ��Ӳ�����¼����ȡ����Դ�ĵĻ�����Ϣ
    strSql = "Select ID,����ID,Ӥ��,������,����ʱ�� From ���Ӳ�����¼ Where Id = [1]"
    Set rs������¼ = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Դ�Ļ�����Ϣ", lngReportID)
    If rs������¼.RecordCount = 0 Then
        Exit Function
    End If
    
    '�ӵ��Ӳ�����������ȡ����Դ�ĵ�������Ϣ
    strSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ʼ�� = [2]  "
    Set rs�������� = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Դ��������Ϣ", lngReportID, intǩ���汾)
    If rs��������.RecordCount = 0 Then
        Exit Function
    End If
    
    If int��ȡ���� = 1 Then
        'ǩ�������ǩ�������Ƿ����
        If thisSign Is Nothing Then
            Exit Function
        End If
    Else
        '��֤ǩ������ǩ����¼����ȡҽ��������ǩ������ǩ��ʱ����Ϣ,ǩ������
        strSql = "Select �����ı� as ҽ������ ,Ҫ�ر�ʾ  as ǩ������ ,�������� From ���Ӳ������� Where �ļ�ID = [1] And �������� = 8 and ��ʼ�� =[2] "
        Set rsǩ����¼ = zlDatabase.OpenSQLRecord(strSql, "��ȡ��󱨸�Դ��ǩ����Ϣ", lngReportID, intǩ���汾)
        If rsǩ����¼.RecordCount = 0 Then
            Exit Function
        End If
        
        '��ȡ��ʽ����ǩ��ʱ�䣬ǩ������
        arr�������� = Split(rsǩ����¼!��������, ";")
        If UBound(arr��������) >= 5 Then
            intRule = Val(arr��������(1))
            strǩ��ʱ�� = Format(arr��������(4), "yyyy-MM-dd HH:mm:ss")
        End If
        If intRule = 0 Then Exit Function
    End If
    
    '���ݹ�����֯����Դ�ģ� ID������ID��Ӥ���������ˣ�����ʱ�䣬ҽ��������ǩ������ǩ��ʱ��,���������������������
    If intRule = 1 Then
        'Դ�Ļ�����Ϣ
        strSourceOut = rs������¼!ID
        strSourceOut = strSourceOut & vbTab & Nvl(rs������¼!����ID)
        strSourceOut = strSourceOut & vbTab & Nvl(rs������¼!Ӥ��)
        strSourceOut = strSourceOut & vbTab & Nvl(rs������¼!������)
        strSourceOut = strSourceOut & vbTab & Nvl(rs������¼!����ʱ��)
        
        'Դ��ǩ����Ϣ
        If int��ȡ���� = 1 Then
            'ǩ������ǩ��������ȡ
            strSourceOut = strSourceOut & vbTab & thisSign.����
            strSourceOut = strSourceOut & vbTab & thisSign.ǩ������
            strSourceOut = strSourceOut & vbTab & Format(thisSign.ǩ��ʱ��, "yyyy-MM-dd HH:mm:ss")
        Else
            '��֤ǩ���������ݿ�ǩ����¼��ȡ
            strSourceOut = strSourceOut & vbTab & Nvl(rsǩ����¼!ҽ������)
            strSourceOut = strSourceOut & vbTab & Nvl(rsǩ����¼!ǩ������)
            strSourceOut = strSourceOut & vbTab & strǩ��ʱ��
        End If
        
        'Դ�ı�������
        rs��������.Filter = "���� ='" & ReportViewType_������� & "'"
        If rs��������.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & Nvl(rs��������!����)
        End If
        
        rs��������.Filter = "���� ='" & ReportViewType_������ & "'"
        If rs��������.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & Nvl(rs��������!����)
        End If
        
        rs��������.Filter = "���� ='" & ReportViewType_���� & "'"
        If rs��������.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & Nvl(rs��������!����)
        End If
    End If
    
    GetSignSourceString = intRule
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
