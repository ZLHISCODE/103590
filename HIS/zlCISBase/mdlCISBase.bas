Attribute VB_Name = "mdlCISBase"
Option Explicit
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrProductName As String
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gblnCancel As Boolean                '��¼�����е�ȡ����ť�Ƿ񱻵����

Public gstrDBOwner As String                '��ǰϵͳ������
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������
Public gstrItemName As String

Public gstrUnitName As String               '�û���λ����
Public gfrmMain As Object

Public gstrMatchMethod As String            'ƥ�䷽ʽ:0��ʾ˫��ƥ��

Public gstrSql As String
Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����
Public gblnOK As Boolean


Public glngPreHWnd As Long '����֧�������ֹ���

Public gobjKernel As New clsCISKernel       '�ٴ����Ĳ���
Public gobjLogisticPlatform As Object       '����ƽ̨�ӿ�
Public gstrPriceClass As String         '�۸�ȼ�

Public gobjRIS As Object                    '����RIS�ӿڶ���
Public Enum RISBaseItemOper                 '����RIS�������ݲ������ͣ�1-������2-�޸ģ�3-ɾ��
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '����RIS�����������ͣ�1��������ĿĿ¼��2��������Ŀ��λ
    ClinicItem = 1
    ClinicItemPart = 2
End Enum

Public gblnKSSStrict As Boolean             '�Ƿ����ÿ���ҩ���ϸ����
Public gblnIncomeItem As Boolean            '��¼������Ŀ�Ƿ�����

Public Type type_user_Digits
    dig_�ɱ��� As Double
    dig_���ۼ� As Double
    dig_���� As Double
    dig_��� As Double
End Type
Public gtype_MaxDigits As type_user_Digits  '������¼��󾫶�

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO
Public Const gstrLisHelp As String = "zl9LisWork"               'LIS���ð���ʱʹ�õĲ�����
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const GCST_INVALIDCHAR = "'"             '�����������Ч�ַ�

'֧�ֻ��ֵĳ���
Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = -4

Public Const GWL_STYLE = (-16)
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

'˽�С�����ģ�����
Public Enum ����_ҩƷĿ¼����_����
    P1_����ҩ������Ŀ = 1
    P2_�г�ҩ������Ŀ = 2
    P3_�в�ҩ������Ŀ = 3
    P4_Ӧ�÷�Χ = 4
    P5_ʱ��ҩƷ�����ε��� = 5
End Enum
Public grsPriceGrade As ADODB.Recordset
 
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Function Select����ѡ����(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional bln����Ա As Boolean = False, _
    Optional strSql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '����:����ѡ����
    '����:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     bln����Ա-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim strPa As String
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "����ѡ����"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    strPa = zlDatabase.GetPara(44, glngSys, 0): strPa = IIf(strPa = "", "11", strPa)
    
    If strSql <> "" Then
    
        gstrSql = strSql
    Else
        gstrSql = "" & _
        "   Select distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
        "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
    
        If str�������� = "" And bln����Ա = False Then
            gstrSql = gstrSql & vbCrLf & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSql = gstrSql & vbCrLf & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c" & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.���� in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) ") & _
            "         AND a.id = c.����id " & _
            IIf(bln����Ա = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSql = gstrSql & vbCrLf & _
            "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) And (a.վ��=[4] or a.վ�� is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            If Mid(strPa, 1, 1) = "1" Then strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlStr.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            If Mid(strPa, 2, 1) = "1" Then strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlStr.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strSql = "" Then
        gstrSql = gstrSql & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSql = gstrSql & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strSql = "" Then
        '�����¼�
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSql, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSql, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.ID, str��������, strKey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "û�����������Ĳ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!ID) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            MsgBox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
        objCtl.Tag = Val(rsTemp!ID)
    End If
    zlCommFun.PressKey vbKeyTab
    Select����ѡ���� = True
End Function

Public Function CheckPriceAdjust(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, Optional ByVal bln�������۹������� As Boolean = False) As Boolean
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '�޿��ʱ���ɱ���ȡҩƷ���ĳɱ���
    '������lngҩƷid-ҩƷ���ID��Ϊ0��������ҩƷ��lng�ⷿid-��Ӧ�ĿⷿID��Ϊ0�������пⷿ��lng����-��Ӧ�����Σ��������-1�򲻹�������
    '      bln�������۹������ԣ�true-������������(���������޸�����ʱ�������޸ĵ�δʵ�ʱ���)
    '���أ�True-������false-�в��������۹���Ҫ���ҩƷ
    '
    Dim rsData As ADODB.Recordset
    Dim str���� As String
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹����򲻽��к�����飬����true
    If Val(zlDatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjust = True: Exit Function
    
    '������޿��
    If lngҩƷID > 0 Then
        If lng�ⷿID > 0 Then
            gstrSql = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] and �ⷿid=[2] " & _
                " And Not (���� = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
            
            If lng���� > 0 Then
                gstrSql = gstrSql & " and Nvl(����,0)=[3] "
            End If
        Else
            gstrSql = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] " & _
                " And Not (���� = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
        End If
        Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "CheckPriceAdjust", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            '�޿��ʱ�����շѼ�Ŀȡ�ۼۣ���ҩƷ���ȡ�ɱ���
            gstrSql = "Select a.�ɱ���, b.�ּ� As �ۼ� " & _
                " From ҩƷ��� A, �շѼ�Ŀ B " & _
                " Where a.ҩƷid = b.�շ�ϸĿid And (Sysdate Between b.ִ������ And b.��ֹ����) " & IIf(bln�������۹������� = False, " And Nvl(a.�Ƿ����۹���, 0) = 1 ", "") & _
                " And b.�ּ� <> a.�ɱ��� And a.ҩƷid = [1] " & GetPriceClassString("B")
            Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "CheckPriceAdjust", lngҩƷID)
            
            If rsData.EOF Then
                'û�ҵ���ʾ�۸�һ��
                CheckPriceAdjust = True
            Else
                '�ҵ���ʾ�۸�һ��
                CheckPriceAdjust = False
            End If
            
            Exit Function
        End If
    End If
    
    If lngҩƷID > 0 Then
        str���� = IIf(str���� = "", "", str����) & " and a.ҩƷid=[1] "
    End If
    
    If lng�ⷿID > 0 Then
        str���� = IIf(str���� = "", "", str����) & " and d.�ⷿid=[2] "
    End If
    
    If lng���� >= 0 Then
        str���� = IIf(str���� = "", "", str����) & " and nvl(d.����,0)=[3] "
    End If
    
    If bln�������۹������� = False Then
        str���� = IIf(str���� = "", "", str����) & " And Nvl(a.�Ƿ����۹���, 0) = 1 "
    End If
    
    gstrSql = "Select ҩƷid, ͨ����, ���, 0 As �ⷿid, '' As �ⷿ, ������, '' As ����, ����, ��λ, ҩ���װ, �ۼ�, Sum(�ɱ��� * ʵ������) / Sum(ʵ������) As �ɱ���, �Ƿ�ʱ��" & vbNewLine & _
        " From (Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ����, c.���, c.���� As ������, Null As ����, a.ҩ�ⵥλ As ��λ, a.ҩ���װ, b.�ּ� As �ۼ�," & vbNewLine & _
        "              nvl(d.ƽ���ɱ���,a.�ɱ���) As �ɱ���, 0 As �Ƿ�ʱ��, d.ʵ������" & vbNewLine & _
        "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
        "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
        "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0  And" & vbNewLine & _
        "             b.�ּ� <> nvl(d.ƽ���ɱ���,a.�ɱ���) " & str���� & GetPriceClassString("B") & vbNewLine & _
        "  And Not (D.���� = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0))" & vbNewLine & _
        " Group By ҩƷid, ͨ����, ���, ������, ����, ��λ, ҩ���װ, �ۼ�, �Ƿ�ʱ�� " & vbNewLine & _
        " Having Sum(ʵ������) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ����, c.���, d.�ⷿid, e.���� As �ⷿ, d.�ϴβ��� As ������, d.�ϴ����� As ����, d.����," & vbNewLine & _
        "       a.ҩ�ⵥλ As ��λ, a.ҩ���װ, d.���ۼ� As �ۼ�, nvl(d.ƽ���ɱ���,a.�ɱ���) As �ɱ���, 1 As �Ƿ�ʱ��" & vbNewLine & _
        " From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
        " Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And nvl(d.���ۼ�,0) <> nvl(d.ƽ���ɱ���,a.�ɱ���)" & vbNewLine & _
        " " & str���� & "" & vbNewLine & _
        "  And Not (D.���� = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) " & vbNewLine & _
        " Order By ͨ����,�ⷿid,����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "CheckPriceAdjust", lngҩƷID, lng�ⷿID, lng����)
    
    'û�ҵ����������۹���Ҫ��ļ�¼������true
    If rsData.EOF Then CheckPriceAdjust = True: Exit Function
    
    CheckPriceAdjust = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'���ܣ���ʼ�������ӿڲ���
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
    End If
End Sub
Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Sub GetMaxDigit()
    '����ȡҩƷ�ĸ�����󾫶�
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSql = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��󾫶�")
    If rsTemp.RecordCount = 0 Then
        gtype_MaxDigits.dig_�ɱ��� = 7
        gtype_MaxDigits.dig_��� = 2
        gtype_MaxDigits.dig_���ۼ� = 7
        gtype_MaxDigits.dig_���� = 7
    Else
        gtype_MaxDigits.dig_�ɱ��� = rsTemp.Fields(1).NumericScale
        gtype_MaxDigits.dig_��� = rsTemp.Fields(0).NumericScale
        gtype_MaxDigits.dig_���ۼ� = rsTemp.Fields(2).NumericScale
        gtype_MaxDigits.dig_���� = rsTemp.Fields(3).NumericScale
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

'ȡҩƷ���۸��������С��λ��
Public Function GetDigit(ByVal int��� As Integer, ByVal int���� As Integer, Optional ByVal int��λ As Integer) As Integer
    'int���1-ҩƷ;2-����
    'int���ݣ�1-�ɱ���;2-���ۼ�;3-����;4-���
    'int��λ�������ȡ���λ�������Բ�����ò���
    '         ҩƷ��λ:1-�ۼ�;2-����;3-סԺ;4-ҩ��;
    '         ���ĵ�λ:1-ɢװ;2-��װ
    '���أ���С2�����Ϊ���ݿ����С��λ��
    
    Dim rsTmp As ADODB.Recordset
    Dim intMax��� As Integer
    Dim intMax�ɱ��� As Integer
    Dim intMax���ۼ� As Integer
    Dim intMax���� As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSql = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, "ȡҩƷ����")
    
    intMax��� = rs.Fields(0).NumericScale
    intMax�ɱ��� = rs.Fields(1).NumericScale
    intMax���ۼ� = rs.Fields(2).NumericScale
    intMax���� = rs.Fields(3).NumericScale
    
    gstrSql = "Select Nvl(����, 0) ���� From ҩƷ���ľ��� Where ��� = [1] And ���� = [2] And ��λ = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ȡҩƷ" & Choose(int����, "�ɱ���", "���ۼ�", "����") & "С��λ��", int���, int����, int��λ)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!����
    End If
    
    If GetDigit = 0 Then
        '���û�����þ��ȣ���ȡ���ݿ���������λ��
        GetDigit = Choose(int����, intMax�ɱ���, intMax���ۼ�, intMax����)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int����, intMax�ɱ���, intMax���ۼ�, intMax����, intMax���)
End Function


Public Function GetUserInfo() As Boolean
    '���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        gstrUserName = UserInfo.����
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MoveSpecialChar(ByVal strInputString As String, Optional ByVal blnMoveSpace As Boolean = True) As String
    '1 ȥ��һ���ַ�: " '_%?"����_%?ת��Ϊ��Ӧ��ȫ���ַ�
    '2 ȥ�������ַ�:�˸��Ʊ����С��س�
    '3 blnMoveSpace���Ƿ�ȥ���ַ��еĿո�Ture-ȥ���ո�ע��ͷβ�ո�Ĭ��ȥ��
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intAsc As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '����ת�����ַ�
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "��"
                Case "%"
                    strTmp = strTmp & "��"
                Case "_"
                    strTmp = strTmp & "��"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intAsc = Asc(Mid(strText, n, 1))
        Select Case intAsc
            Case 8, 9, 10, 13
            Case 32
                '�ո���
                If blnMoveSpace = False Then
                    strTmp = strTmp & Mid(strText, n, 1)
                End If
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function

Public Function zlClinicCodeRepeat(strInputCode As String, Optional lngSelfID As Long) As Boolean
    '----------------------------------
    '���ܣ����������Ŀ������Ƿ������б����ظ����ظ��������ʾ
    '��Σ�strInputCode-����ı��룻lngSelfID-�Լ���ID�ţ����޸�ʱ����Ҫ��������������ж�
    '���Σ��ظ�����True��������Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.����||' ['||I.����||']'||I.���� as ����" & _
            " from ������ĿĿ¼ I,������Ŀ��� K" & _
            " where I.���=K.���� and I.����=[1] " & _
            "       and I.ID<>[2]"
    err = 0: On Error GoTo ErrHand
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", strInputCode, lngSelfID)
        
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "����Ŀ�롰" & !���� & "�������ظ���", vbExclamation, gstrSysName
            zlClinicCodeRepeat = True
        Else
            zlClinicCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlClinicCodeRepeat = True
End Function


Public Function zlExistItem(ByVal strTbleName As String, ByVal strField As String, ByVal varValues As Variant, _
                            ByVal strItemName As String) As Boolean
    
    '----------------------------------
    '���ܣ������Ŀ�Ƿ����,���ڲ�������ʱ�ļ��
    '��Σ�strTableName ���� ,strField �ֶ��� , ,lngItemID,�ֶε�ֵ,strItemName ��ʾʱ��ʾ����Ŀ����
    '���Σ����ڷ���True��������Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    err = 0: On Error GoTo ErrHand
    strSql = "Select " & strField & " From " & strTbleName & " Where " & strField & "=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", varValues)
    If rsTmp.RecordCount > 0 Then
        zlExistItem = True
    Else
         MsgBox "��" & strItemName & "���Ѿ�����������Աɾ����", vbExclamation, gstrSysName
        zlExistItem = False
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlExistItem = False
End Function

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    X = objPoint.X * 15 + objBill.CellLeft
    Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intType As Integer
    Dim curDate As Date
    
    On Error GoTo errH
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    strSql = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!��Ź���)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSql = Format(CDate(Format(rsTmp!����, "YYYY-MM-dd")) - CDate(Format(rsTmp!����, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & strSql & Format(Right(strNo, 4), "0000")
    Else
        '������
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DelInvalidChar(ByVal strchar As String, Optional ByVal strInvalidChar As String) As String
    'ɾ���Ƿ��ַ�
    'strChar: Ҫ������ַ�
    'strInvalidChar���Ƿ��ַ��������Ϊ�գ���Ϊ~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,���򰴴�����ַ�����
    Dim strBit As String, i As Integer, strWord As String
    strWord = "~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>"
    If strInvalidChar <> "" Then strWord = strInvalidChar
    If Len(strchar) > 0 Then
        For i = 1 To Len(strchar)
            strBit = Mid$(strchar, i, 1)
            If InStr(strWord, strBit) <= 0 Then
                DelInvalidChar = DelInvalidChar & strBit
            End If
        Next
    End If
End Function

Public Function CheckKSSPrivilege() As Boolean
'���ܣ����ϵͳ�Ƿ���ڿ���ҩ����Ȩ����Ա���������õ�ǰ����Ա����ҩ����UserInfo����
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    UserInfo.��ҩ���� = 0
    
    On Error GoTo errH
    strSql = "Select ���� From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and ��ԱID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        UserInfo.��ҩ���� = Val("" & rsTmp!����)
        CheckKSSPrivilege = True
    Else
        strSql = "Select 1 From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and Rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel")
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub GetPriceClass()
    '���ݵ�¼վ���ȡҩƷ�ļ۸�ȼ�
    Dim rsData As ADODB.Recordset
    
    If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
        gstrSql = " Select a.�۸�ȼ� " & _
            " From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B " & _
            " Where a.�۸�ȼ� = b.���� And a.���� = 0 And b.�Ƿ�����ҩƷ = 1 And a.վ�� = [1] And Nvl(b.����ʱ��, Sysdate + 1) > Sysdate "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "GetPriceClass", gstrNodeNo)
        
        If rsData.RecordCount > 0 Then gstrPriceClass = rsData!�۸�ȼ�
    End If
End Sub

Public Function GetPriceClassString(strTableName As String) As String
    '���ݴ����ı������ؼ۸�ȼ���������
    GetPriceClassString = " And " & IIf(strTableName = "", "�۸�ȼ� Is Null ", strTableName & ".�۸�ȼ� Is Null ")
    
End Function

Public Function zlGetrsPriceGrade(ByRef rsOutPriceGrade As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�۸�ȼ���¼��
    '���:
    '����:rsOutPriceGrade-���ؼ۸�ȼ���δ���û��ȡʧ����ʱ������Nothing
    '����:�����ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2017-06-30 14:08:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    If Not grsPriceGrade Is Nothing Then
        If grsPriceGrade.State = 1 Then Set rsOutPriceGrade = grsPriceGrade: zlGetrsPriceGrade = True: Exit Function
    End If
    '����Ƿ����ã��������
    strSql = "" & _
    "   Select ����,���� From �շѼ۸�ȼ� A where nvl(����ʱ��,sysdate+1)>sysdate Order by ����"
    Set grsPriceGrade = zlDatabase.OpenSQLRecord(strSql, "��ȡ�۸�ȼ�")
    Set rsOutPriceGrade = grsPriceGrade
    zlGetrsPriceGrade = rsOutPriceGrade.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function FmgFlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'֧��frmDoctorManage������ֵĹ���
    On Error GoTo errH
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
            Case -7864320  '���¹�
                If frmDoctorManage.vscBar.Value <> frmDoctorManage.vscBar.Max Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageDown
                End If
            Case 7864320   '���Ϲ�
                If frmDoctorManage.vscBar.Value <> 0 Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageUp
                End If
        End Select
    End Select
    FmgFlexScroll = CallWindowProc(glngPreHWnd, hWnd, wMsg, wParam, lParam)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowSpecChar(frmParent As Object) As String
'���ܣ���ģ̬�������������ַ�����
'������frmParent=���ø�����
'���أ�ѡ��������ַ�����ȡ���������ؿ�
    Dim frmNew As frmSpecChar
    Set frmNew = New frmSpecChar
    frmNew.Show 1, frmParent
    If gblnOK Then ShowSpecChar = frmNew.mstrChar
End Function

Public Sub ArrayIcons(objLvw As ListView, Optional intBegin As Integer = 1, Optional blnShow As Boolean)
'���ܣ����ݵ�һ��ͼ���λ��������������ͼ��
    Dim i As Integer, t As Long
    Dim r As RECT

    Call GetClientRect(objLvw.hWnd, r)
    
    If blnShow Then
        If objLvw.ListItems(intBegin).Top < 30 Then
           objLvw.ListItems(intBegin).Top = 30
        ElseIf objLvw.ListItems(intBegin).Top + objLvw.ListItems(intBegin).Height > (r.Bottom - r.Top) * Screen.TwipsPerPixelY Then
            objLvw.ListItems(intBegin).Top = (r.Bottom - r.Top) * Screen.TwipsPerPixelY - objLvw.ListItems(intBegin).Height
        End If
    End If
    
    '�����ͼ��
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            'Item��Width�������ֲ���,Left��ָͼ��
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t + .Height
        End With
    Next
    
    '�����ͼ��
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To 1 Step -1
        With objLvw.ListItems(i)
            'Item��Width�������ֲ���,Left��ָͼ��
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t - .Height
        End With
    Next
End Sub
