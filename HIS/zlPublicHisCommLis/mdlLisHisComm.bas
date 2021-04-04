Attribute VB_Name = "mdlLisHisComm"
Option Explicit
Public gblnInit As Boolean                                         '���������Ƿ��ѳ�ʼ��

Public grsParas As ADODB.Recordset                                  'ϵͳ��������
Public grsUserParas As ADODB.Recordset                              'ϵͳ��������
Public gcolPrivs As Collection                                      '��ǰ�û��߱������г���Ĺ���Ȩ��
Public gblnAllSite As Boolean                                       '�Ƿ��ܹ��鿴����վ��
  
Public gobjHisComLib As Object
Public gobjHisDatabase As Object
Public gobjHisSystem As Object
Public gobjPlugIn As Object                                         '��Ҳ���

Public gcnLisOracle As New ADODB.Connection                         'LIS�������ݿ�����
Public gcnHisOracle As New ADODB.Connection                         'HIS�������ݿ�����

Public gstrDBUser As String

Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

Public Type TYPE_SYS_INFO   '-----------Ӧ�ó�����Ϣ �� ע����Ϣ
    AppName As String       'ϵͳ���� (��Ʒ���+����������������ҽҵ���)
    ShortName As String     '��Ʒ����
    AppTitle As String      'ϵͳ���⣬��Ʒȫ��
    
    Version As String       'ϵͳ�汾
    AviPath As String       'AVI�ļ�·��
    
    UnitName  As String     '�û���λ����
    Supporter As String     '����֧����
    Develop As String       '������
    SupporterWEB As String  '֧����WEB����
    SupporterMail As String '֧�����ʼ�
    SupporterURL As String  '֧������ַ
    ProductLine  As String  '��Ʒϵ�У�[��׼��],[��ͻ���]
    
    SysNo       As Long     'ϵͳ���
    ModlNo      As Long     'ģ���
    VersionHIS As String       'HISϵͳ�汾
    VersionLIS As String       'lisϵͳ�汾
End Type

Public Type TYPE_SYS_PARAMETER    'ϵͳ����
    Privs        As String  'ģ��Ȩ��
    MachineCount As Integer '��������
    blnEmerge    As Boolean '�Ƿ����ּ���
    BuffDir      As String   '���ػ����¼���Ļ���Ŀ¼
    InvaidWord   As String   '��ȥ���ķǳ��ַ�
    intCA        As Integer  'CA���ı��
    strMatch     As String   '����ƥ��
End Type

Public Type TYPE_USER_INFO
    ID As Long          '��ԱID
    DeptID As Long      '��Ա��Ӧ�Ĳ���ID
    DeptName As String  '��Ա��Ӧ�Ĳ�������
    No As String        '��Ա���
    Name As String      '��Ա����
    Code As String      '��Ա����
    DBUser As String    '��Ա��Ӧ�����ݿ��û���
    ComputerName As String          '������
    NodeNo As String                '��Ա��½վ��
End Type

Public gSampleShowColour As SampleValShowColour                    '�����ʾ��ɫ
Public Type SampleValShowColour                                    '�����ɫ��ʾ
    ���� As Double
    ƫ�� As Double
    ƫ�� As Double
    �쳣 As Double
    ��ʾƫ�� As Double
    ��ʾƫ�� As Double
    ����ƫ�� As Double
    ����ƫ�� As Double
End Type


Public gUserInfo As TYPE_USER_INFO
Public gSysInfo As TYPE_SYS_INFO
Public gSysParameter As TYPE_SYS_PARAMETER
Public gstrSQL  As String

Private Const pҽ�����ѹ��� As Integer = 1257                       '���˷���ģ����Ȩ
Private Const p����ҽ���´� As Integer = 1252                       '����ҽ���´�
Private Const pסԺҽ���´� As Integer = 1253                       'סԺҽ���´�
Private Const p���ﲡ������ As Integer = 1250                       '���ﲡ��
Private Const pסԺ�������� As Integer = 1251



Public Declare Function SetParent Lib "user32.dll " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const CB_ADDSTRING = &H143
Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private mstrPara As String

Public Function ComInitComLib(ByRef strErr As String) As Boolean
'��ʼ����������,�ڳ�������ʱ����
    Dim strDBUser As String
    On Error GoTo errH
    ComInitComLib = False



    If gblnInit Then
        ComInitComLib = True
        Exit Function
    End If

    If gcnHisOracle.State = 1 Then
        Set gobjHisComLib = CreateObject("zl9ComLib.clsComLib")
        gobjHisComLib.InitCommon gcnHisOracle
        Set gobjHisDatabase = gobjHisComLib.zlDatabase
        If VerCompare(gSysInfo.VersionHIS, "10.35.10") = -1 Then
            strDBUser = GetUserDB(2)
            gobjHisComLib.SetDbUser strDBUser
        End If
        Set gobjHisSystem = CreateObject("zl9ComLib.clsSystem")
    End If


    ComInitComLib = True
    gblnInit = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
End Function

'���� ��������ComLib��һЩ������������
Public Function ComOpenSQL(ByVal selDB As Integer, ByVal strSQL As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
    '���ܣ�ͨ��ComLib����򿪴�����SQL�ļ�¼��
    '
    Dim lngCount As Long
    Dim var(30) As Variant
    

    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        Err.Raise -2147483645, , "��֧�ֳ���30��������SQL��"
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    
    Set ComOpenSQL = OpenSQLRecord(selDB, strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))



End Function

Public Function ComExecuteProc(ByVal selDB As Integer, strSQL As String, ByVal strFormCaption As String) As String
    '���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
    '���أ��޴��󷵻ؿմ������򷵻ش�����ʾ
    

    Call ExecuteProcedure(selDB, strSQL, strFormCaption)
    
End Function

Public Function BeforCreateLisValueStr(ByVal strAdvices As String, Optional ByVal DateE As Date, Optional strErr As String) As Boolean
          '����ҽ��ID��ʱ��,�жϴ���ʱ��֮���Ƿ����ҽ��ID��Ӧ������˼�¼
          'stradvices     'ҽ��ID,���ҽ��ʹ��","�ŷָ�
          'DateE          '��ȡ�ڸ�ʱ��֮���Ƿ����ҽ��ID��Ӧ�ļ�¼
          
          '����       True=���ڼ�¼,False=�����ڼ�¼
          
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo BeforCreateLisValueStr_Error

2         BeforCreateLisValueStr = False
3         strErr = ""
          
4         strSQL = "Select /*+cardinality(c,10)*/ B.ID From ����������� A, ���鱨���¼ B,Table(f_Str2list([1])) C" & _
                   " Where a.�걾id = b.Id And a.����id=c.Column_Value and b.���ʱ�� > [2]"
5         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����Ƿ�������", strAdvices, DateE)
          
          '�м�¼ʱ����True
6         Do While rsTmp.EOF = False
7             If IsNull(rsTmp("ID")) = False Then
8                 BeforCreateLisValueStr = True
9                 Exit Function
10            End If
11            rsTmp.MoveNext
12        Loop


13        Exit Function
BeforCreateLisValueStr_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(BeforCreateLisValueStr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear
              
End Function

Public Function CreateLisValueStr(strAdvices As String, Optional lngPatient As Long, Optional strErr As String, Optional intType As Integer) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                  ���ݴ���ҽ��ID���ؽ��
          '����
          '           strAdvices          ����ID��,�ö��ŷָ�
          '           lngPatient          ��ѡ�Ĳ��������벡��ID��ֻ������ID���ҽ��
          '           strType               0-��ˣ�1-ȡ�����
          '�걾��ɸ�ʽ
          '               ����(1=��ͨ)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2> Ӥ����� <split2>
          '                   ָ��1<split4>������1<split4>��λ1<split4>�����־1<split4>�������1<split4>�������1<split4>��˽��Ŀ1<split4>ָ�����1<split4>������1<split4>Ӣ����1<split4>�ο���ֵ1<split4>�ο���ֵ1<split4>С��λ��1<split3>
          '                   ָ��2<split4>������2<split4>��λ2<split4>�����־2<split4>�������2<split4>�������2<split4>��˽��Ŀ2<split4>ָ�����2<split4>������2<split4>Ӣ����2<split4>�ο���ֵ2<split4>�ο���ֵ2<split4>С��λ��2<split3>
          '                   ָ��3<split4>������3<split4>��λ3<split4>�����־3<split4>�������3<split4>�������3<split4>��˽��Ŀ3<split4>ָ�����3<split4>������3<split4>Ӣ����3<split4>�ο���ֵ3<split4>�ο���ֵ3<split4>С��λ��3<split1>
          '
          '               ����(2=΢����)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>
          '               ϸ����1<split3>����1<split3>��ҩ����1<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split2>
          '               ϸ����2<split3>����2<split3>��ҩ����2<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split1>
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim lngID As Long
          Dim lngSampleID As Long
          Dim lngSampleGroup As Long
          Dim lngMicroID As Long
          Dim strSampleOne As String
          Dim strSampleTwo As String
          Dim intCount As Integer
          Dim strTemp As String
          Dim stridSQL As String
          Dim i As Long
          Dim str�ο���ֵ As String
          Dim str�ο���ֵ As String
          Dim str����ο� As String
          Dim str������ As String
          
          '�ָ��ĳ���
          Const conSplit1 As String = "<split1>"                        '���ڷָ��걾,ʹ�á�<split1>���ָ�����ǰʹ�á�|��
          Const conSplit2 As String = "<split2>"                        '���ڷָ��걾��Ϣ,ʹ�á�<split2>���ָ�����ǰʹ�á�;��
          Const conSplit3 As String = "<split3>"                        '���ڷָ��걾ָ����Ϣ,ʹ�á�<split3>���ָ�����ǰʹ�á�,��
          Const conSplit4 As String = "<split4>"                        '���ڷָ�ָ������Ϣ,ʹ�á�<split4>���ָ�����ǰʹ�á�^��
          
          
          
          '�ֱ������ͨ��΢������Ŀ
          
          '��strAdvices�����ַ������ȳ���4000������
1         On Error GoTo CreateLisValueStr_Error

2         If Len(strAdvices) > 4000 Then
3             For i = 0 To UBound(Split(strAdvices, ","))
4                 strTemp = strTemp & "," & Split(strAdvices, ",")(i)
5                 intCount = intCount + 1
                  
6                 If intCount = 200 Then
7                     stridSQL = stridSQL & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
8                     intCount = 0
9                     strTemp = ""
10                End If
11            Next
12            If strTemp <> "" Then
13                stridSQL = Mid(stridSQL, 12) & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
14            End If
15        Else
16            stridSQL = "Select Column_Value From Table(Cast(F_Num2list([3]) As Zltools.T_Numlist))"
17        End If
          
          '��ͨ
18        strSQL = "Select distinct [�ؼ���] 1 Type,A.����id,a.������Դ ,b.����ʱ��,b.������,b.�����,b.���ʱ��,e.���� ������Ŀ����, " & vbNewLine & _
                  "       D.������ || '(' || D.Ӣ���� || ')' ָ��,c.������, D.��λ," & vbNewLine & _
                  "       Decode(C.�����־, 1, '', 2, '��', 3, '��', 4, '�쳣', 5, '����', 6, '����', '') �����־," & vbNewLine & _
                  "       c.����ο�,C.�������, 0 ��˽��Ŀ,b.�걾����,a.Ӥ�� Ӥ�����,d.ָ�����,a.�걾id,a.���id,d.������,d.Ӣ����,c.�ο���ֵ,c.�ο���ֵ,nvl(d.С��λ��,2) С��λ��,b.������Դ,d.������� " & vbNewLine & _
                  "From ����������� A, ���鱨���¼ B, ���鱨����ϸ C, ����ָ�� D,���������Ŀ e [���]" & vbNewLine & _
                  "Where A.�걾id = B.Id And B.Id = C.�걾id And C.��Ŀid = D.Id And Nvl(B.΢����, 0) <> 1 And" & vbNewLine & _
                  "      a.���id = c.���id and a.���id = e.id  [����1]" & vbNewLine & _
                  "     [����] " & vbNewLine & _
                  " order by a.����id,a.�걾ID,a.���id,c.������� "
19        If intType = 1 Then
20           strSQL = Replace(strSQL, "[����1]", "")
21        Else
22            strSQL = Replace(strSQL, "[����1]", " and   b.����� is not null")
23        End If
24        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[����]", " and A.����id In (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) ")
      '        strSQL = Replace(strSQL, "[����]", " and A.����id In (" & stridSQL & ")")
25            strSQL = Replace(strSQL, "[�ؼ���]", "/*+cardinality(f,10)*/")
26            strSQL = Replace(strSQL, "[���]", ",(" & stridSQL & ") f")
27            strSQL = Replace(strSQL, "[����]", " and a.����id=f.Column_Value")
28        Else
29            strSQL = Replace(strSQL, "[�ؼ���]", "")
30            strSQL = Replace(strSQL, "[���]", "")
31            strSQL = Replace(strSQL, "[����]", " and A.HIS����ID = [1] ")
32        End If

33        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatient, Mid(strTemp, 2), strAdvices)
34        lngID = 0
35        Do Until rsTmp.EOF
              '���ݸ�ʽ���
36            str�ο���ֵ = IIf(IsNumeric(rsTmp("�ο���ֵ") & ""), IIf(Mid(rsTmp("�ο���ֵ") & "", 1, 1) = ".", "0" & rsTmp("�ο���ֵ"), rsTmp("�ο���ֵ") & ""), rsTmp("�ο���ֵ") & "")
37            str�ο���ֵ = IIf(IsNumeric(rsTmp("�ο���ֵ") & ""), IIf(Mid(rsTmp("�ο���ֵ") & "", 1, 1) = ".", "0" & rsTmp("�ο���ֵ"), rsTmp("�ο���ֵ") & ""), rsTmp("�ο���ֵ") & "")
              If IsNumeric(rsTmp("������") & "") Then
38               str������ = IIf(Val(rsTmp("�������") & "") = 1, Format(rsTmp("������") & "", IIf(Val(rsTmp("С��λ��") & "") > 0, "0." & String(Val(rsTmp("С��λ��") & ""), "0"), "0")), rsTmp("������") & "")
              Else
                  str������ = rsTmp("������") & ""
              End If
39            If InStr(rsTmp("����ο�") & "", "--") > 0 Then
40                str����ο� = IIf(IsNumeric(Split(rsTmp("����ο�") & "", "--")(0)), IIf(Mid(Split(rsTmp("����ο�") & "", "--")(0), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "--")(0), Split(rsTmp("����ο�") & "", "--")(0)), Split(rsTmp("����ο�") & "", "--")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("����ο�") & "", "--")(1)), IIf(Mid(Split(rsTmp("����ο�") & "", "--")(1), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "--")(1), Split(rsTmp("����ο�") & "", "--")(1)), Split(rsTmp("����ο�") & "", "--")(1))
41            ElseIf InStr(rsTmp("����ο�") & "", "-") > 0 Then
42                str����ο� = IIf(IsNumeric(Split(rsTmp("����ο�") & "", "-")(0)), IIf(Mid(Split(rsTmp("����ο�") & "", "-")(0), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "-")(0), Split(rsTmp("����ο�") & "", "-")(0)), Split(rsTmp("����ο�") & "", "-")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("����ο�") & "", "-")(1)), IIf(Mid(Split(rsTmp("����ο�") & "", "-")(1), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "-")(1), Split(rsTmp("����ο�") & "", "-")(1)), Split(rsTmp("����ο�") & "", "-")(1))
43            Else
44                str����ο� = rsTmp("����ο�") & ""
45            End If
              
46            If lngID <> NVL(rsTmp("����ID"), 0) Or lngSampleID <> NVL(rsTmp("�걾ID"), 0) Or lngSampleGroup <> NVL(rsTmp("���id"), 0) Then
47                strSampleOne = strSampleOne & conSplit1 & "1" & conSplit2 & rsTmp("����ID") & conSplit2 & rsTmp("������Դ") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("����ʱ��")), "", rsTmp("����ʱ��"))) & conSplit2 & _
                              rsTmp("������") & conSplit2 & rsTmp("�����") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("���ʱ��")), "", rsTmp("���ʱ��"))) & conSplit2 & rsTmp("������Ŀ����") & conSplit2 & _
                              rsTmp("�걾����") & conSplit2 & NVL(rsTmp("Ӥ�����"), "0") & conSplit2 & rsTmp("ָ��") & conSplit4 & str������ & conSplit4 & rsTmp("��λ") & _
                              conSplit4 & rsTmp("�����־") & conSplit4 & str����ο� & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("��˽��Ŀ") & conSplit4 & rsTmp("ָ�����") & _
                              conSplit4 & rsTmp("������") & conSplit4 & rsTmp("Ӣ����") & conSplit4 & str�ο���ֵ & conSplit4 & str�ο���ֵ & IIf(rsTmp("������Դ") & "" = "4", conSplit4 & rsTmp("С��λ��"), "")
48            Else
49                strSampleOne = strSampleOne & conSplit3 & rsTmp("ָ��") & conSplit4 & str������ & conSplit4 & rsTmp("��λ") & _
                              conSplit4 & rsTmp("�����־") & conSplit4 & str����ο� & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("��˽��Ŀ") & conSplit4 & rsTmp("ָ�����") & _
                              conSplit4 & rsTmp("������") & conSplit4 & rsTmp("Ӣ����") & conSplit4 & str�ο���ֵ & conSplit4 & str�ο���ֵ & IIf(rsTmp("������Դ") & "" = "4", conSplit4 & rsTmp("С��λ��"), "")
50            End If
51            lngID = NVL(rsTmp("����ID"), 0)
52            lngSampleID = NVL(rsTmp("�걾ID"), 0)
53            lngSampleGroup = NVL(rsTmp("���id"), 0)
54            rsTmp.MoveNext
55        Loop
          
          
56        lngID = 0
57        lngMicroID = 0
58        strSQL = "Select distinct [�ؼ���] 2 Type,A.����id,a.������Դ ,b.����ʱ��,b.������,b.�����,b.���ʱ��,g.���� ������Ŀ����, " & vbNewLine & _
                  "       E.������ || '(' || E.Ӣ���� || ')' ϸ��, C.��������, C.��ҩ����, F.������ || '(' || F.Ӣ���� || ')' ������, D.��� �����ؽ��," & vbNewLine & _
                  "       D.�������, D.ҩ������, F.�÷�����1, F.�÷�����2, ѪҩŨ��1, ѪҩŨ��2, ��ҩŨ��1, ��ҩŨ��2,c.ϸ��ID,b.�걾����,a.Ӥ�� Ӥ�����" & vbNewLine & _
                  "From ����������� A, ���鱨���¼ B, ���鱨��ϸ�� C, ���鱨��ҩ�� D, ����ϸ����¼ E, ����ҩ�� F,���������Ŀ g [���]" & vbNewLine & _
                  "Where A.�걾id = B.Id And B.΢���� = 1 And B.Id = C.�걾id And C.Id = D.���id And C.ϸ��id = E.Id And D.ҩ��id = F.Id and a.���id = g.id " & vbNewLine & _
                  "      [����] " & vbNewLine & _
                  " order by a.����id,c.ϸ��id"
                  
59        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[����]", " and A.����id In (" & stridSQL & ")")
60            strSQL = Replace(strSQL, "[�ؼ���]", "/*+cardinality(h,10)*/")
61            strSQL = Replace(strSQL, "[���]", ",(" & stridSQL & ") h")
62            strSQL = Replace(strSQL, "[����]", " and a.����id=h.Column_Value")
63        Else
64            strSQL = Replace(strSQL, "[�ؼ���]", "")
65            strSQL = Replace(strSQL, "[���]", "")
66            strSQL = Replace(strSQL, "[����]", " and A.HIS����ID = [1] ")
67        End If
68        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatient, Mid(strTemp, 2), strAdvices)
          
          
          '               ����(2=΢����)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>Ӥ�����<split2>
          '               ϸ����1<split3>����1<split3>��ҩ����1<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split2>
          '               ϸ����2<split3>����2<split3>��ҩ����2<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split1>
          
          
69        If rsTmp.RecordCount <= 0 Then
70            strSQL = "Select  distinct [�ؼ���] 2 Type, a.����id, a.������Դ, b.����ʱ��, b.������, b.�����, b.���ʱ��, c.������ ������Ŀ����, c.δ��� ϸ��, c.��������, c.��ҩ����, '' ������, '' �����ؽ��, '' �������," & vbNewLine & _
                       "          '' ҩ������, '' �÷�����1, '' �÷�����2, '' ѪҩŨ��1, '' ѪҩŨ��2, '' ��ҩŨ��1, '' ��ҩŨ��2, c.ϸ��id, b.�걾����, a.Ӥ�� Ӥ�����" & vbNewLine & _
                       "   From ����������� A, ���鱨���¼ B, ���鱨��ϸ�� C [���]" & vbNewLine & _
                       "   Where a.�걾id = b.Id And b.΢���� = 1 And b.Id = c.�걾id " & vbNewLine & _
                       "      [����] " & vbNewLine & _
                       "   Order By a.����id, c.ϸ��id"
                      
71            If lngPatient = 0 Then
      '            strSQL = Replace(strSQL, "[����]", " and A.����id In (" & stridSQL & ")")
72                strSQL = Replace(strSQL, "[�ؼ���]", "/*+cardinality(d,10)*/")
73                strSQL = Replace(strSQL, "[���]", ",(" & stridSQL & ") D")
74                strSQL = Replace(strSQL, "[����]", " and a.����id=d.Column_Value")
75            Else
76                strSQL = Replace(strSQL, "[�ؼ���]", "")
77                strSQL = Replace(strSQL, "[���]", "")
78                strSQL = Replace(strSQL, "[����]", " and A.HIS����ID = [1] ")
79            End If
80            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatient, Mid(strTemp, 2), strAdvices)
81        End If
          
82        Do Until rsTmp.EOF
83            If lngID <> NVL(rsTmp("����ID"), 0) Then
84                strSampleTwo = strSampleTwo & conSplit1 & "2" & conSplit2 & IIf(IsNull(rsTmp("����ID")), "", rsTmp("����ID")) & conSplit2 & IIf(IsNull(rsTmp("������Դ")), "", rsTmp("������Դ")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("����ʱ��")), "", rsTmp("����ʱ��"))) & conSplit2 & _
                              IIf(IsNull(rsTmp("������")), "", rsTmp("������")) & conSplit2 & IIf(IsNull(rsTmp("�����")), "", rsTmp("�����")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("���ʱ��")), "", rsTmp("���ʱ��"))) & conSplit2 & IIf(IsNull(rsTmp("������Ŀ����")), "", rsTmp("������Ŀ����")) & conSplit2 & _
                              IIf(IsNull(rsTmp("�걾����")), "", rsTmp("�걾����")) & conSplit2 & NVL(rsTmp("Ӥ�����"), "0") & conSplit2 & IIf(IsNull(rsTmp("ϸ��")), "", rsTmp("ϸ��")) & conSplit3 & IIf(IsNull(rsTmp("��������")), "", rsTmp("��������")) & conSplit3 & IIf(IsNull(rsTmp("��ҩ����")), "", rsTmp("��ҩ����")) & _
                              conSplit3 & IIf(IsNull(rsTmp("������")), "", rsTmp("������")) & conSplit4 & IIf(IsNull(rsTmp("�����ؽ��")), "", rsTmp("�����ؽ��")) & conSplit4 & IIf(IsNull(rsTmp("�������")), "", rsTmp("�������")) & conSplit4 & IIf(IsNull(rsTmp("ҩ������")), "", rsTmp("ҩ������")) & conSplit4 & IIf(IsNull(rsTmp("�÷�����1")), "", rsTmp("�÷�����1")) & _
                              conSplit4 & IIf(IsNull(rsTmp("�÷�����2")), "", rsTmp("�÷�����2")) & conSplit4 & IIf(IsNull(rsTmp("ѪҩŨ��1")), "", rsTmp("ѪҩŨ��1")) & conSplit4 & IIf(IsNull(rsTmp("ѪҩŨ��2")), "", rsTmp("ѪҩŨ��2")) & conSplit4 & IIf(IsNull(rsTmp("��ҩŨ��1")), "", rsTmp("��ҩŨ��1")) & conSplit4 & IIf(IsNull(rsTmp("��ҩŨ��2")), "", rsTmp("��ҩŨ��2"))
85                lngMicroID = NVL(rsTmp("ϸ��ID"), 0)
86            Else
87                If lngMicroID <> NVL(rsTmp("ϸ��ID"), 0) Then
88                    strSampleTwo = strSampleTwo & conSplit2 & rsTmp("ϸ��") & conSplit3 & rsTmp("��������") & conSplit3 & rsTmp("��ҩ����") & _
                              conSplit3 & rsTmp("������") & conSplit4 & rsTmp("�����ؽ��") & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("ҩ������") & conSplit4 & rsTmp("�÷�����1") & _
                              conSplit4 & rsTmp("�÷�����2") & conSplit4 & rsTmp("ѪҩŨ��1") & conSplit4 & rsTmp("ѪҩŨ��2") & conSplit4 & rsTmp("��ҩŨ��1") & conSplit4 & rsTmp("��ҩŨ��2")
89                Else
90                    strSampleTwo = strSampleTwo & conSplit3 & rsTmp("������") & conSplit4 & rsTmp("�����ؽ��") & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("ҩ������") & _
                              conSplit4 & rsTmp("�÷�����1") & conSplit4 & rsTmp("�÷�����2") & conSplit4 & rsTmp("ѪҩŨ��1") & conSplit4 & rsTmp("ѪҩŨ��2") & _
                              conSplit4 & rsTmp("��ҩŨ��1") & conSplit4 & rsTmp("��ҩŨ��2")
91                End If
92            End If
93            lngID = NVL(rsTmp("����ID"), 0)
94            lngMicroID = NVL(rsTmp("ϸ��ID"), 0)
95            rsTmp.MoveNext
96        Loop

97        If strSampleOne <> "" Then
98            strSampleOne = Mid(strSampleOne, Len(conSplit1) + 1)
99        End If
100       If strSampleTwo <> "" Then
101           strSampleTwo = Mid(strSampleTwo, Len(conSplit1) + 1)
102       End If
103       If strSampleTwo <> "" Then
104           If strSampleOne = "" Then
105               CreateLisValueStr = strSampleTwo
106           Else
107               CreateLisValueStr = strSampleOne & conSplit1 & strSampleTwo
108           End If
109       Else
110           CreateLisValueStr = strSampleOne
111       End If


112       Exit Function
CreateLisValueStr_Error:
113       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(CreateLisValueStr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
114       Err.Clear
          
End Function

Public Function CreateLisValueStrForTJ(strAdvices As String, Optional lngPatient As Long, Optional strErr As String, Optional intType As Integer) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                  ���ݴ���ҽ��ID���ؽ��
          '����
          '           strAdvices          ����ID��,�ö��ŷָ�
          '           lngPatient          ��ѡ�Ĳ��������벡��ID��ֻ������ID���ҽ��
          '           strType               0-��ˣ�1-ȡ�����
          '�걾��ɸ�ʽ
          '               ����(1=��ͨ)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2> Ӥ����� <split2>
          '                   ָ��1<split4>������1<split4>��λ1<split4>�����־1<split4>�������1<split4>�������1<split4>��˽��Ŀ1<split4>ָ�����1<split4>������1<split4>Ӣ����1<split4>�ο���ֵ1<split4>�ο���ֵ1<split4>С��λ��1<split3>
          '                   ָ��2<split4>������2<split4>��λ2<split4>�����־2<split4>�������2<split4>�������2<split4>��˽��Ŀ2<split4>ָ�����2<split4>������2<split4>Ӣ����2<split4>�ο���ֵ2<split4>�ο���ֵ2<split4>С��λ��2<split3>
          '                   ָ��3<split4>������3<split4>��λ3<split4>�����־3<split4>�������3<split4>�������3<split4>��˽��Ŀ3<split4>ָ�����3<split4>������3<split4>Ӣ����3<split4>�ο���ֵ3<split4>�ο���ֵ3<split4>С��λ��3<split1>
          '
          '               ����(2=΢����)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>
          '               ϸ����1<split3>����1<split3>��ҩ����1<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split2>
          '               ϸ����2<split3>����2<split3>��ҩ����2<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split1>
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim lngID As Long
          Dim lngSampleID As Long
          Dim lngSampleGroup As Long
          Dim lngMicroID As Long
          Dim strSampleOne As String
          Dim strSampleTwo As String
          Dim intCount As Integer
          Dim strTemp As String
          Dim stridSQL As String
          Dim i As Long
          Dim str�ο���ֵ As String
          Dim str�ο���ֵ As String
          Dim str����ο� As String
          Dim str������ As String
          
          '�ָ��ĳ���
          Const conSplit1 As String = "<split1>"                        '���ڷָ��걾,ʹ�á�<split1>���ָ�����ǰʹ�á�|��
          Const conSplit2 As String = "<split2>"                        '���ڷָ��걾��Ϣ,ʹ�á�<split2>���ָ�����ǰʹ�á�;��
          Const conSplit3 As String = "<split3>"                        '���ڷָ��걾ָ����Ϣ,ʹ�á�<split3>���ָ�����ǰʹ�á�,��
          Const conSplit4 As String = "<split4>"                        '���ڷָ�ָ������Ϣ,ʹ�á�<split4>���ָ�����ǰʹ�á�^��
          
          
          
          '�ֱ������ͨ��΢������Ŀ
          
          '��strAdvices�����ַ������ȳ���4000������
1         On Error GoTo CreateLisValueStrForTJ_Error

2         If Len(strAdvices) > 4000 Then
3             For i = 0 To UBound(Split(strAdvices, ","))
4                 strTemp = strTemp & "," & Split(strAdvices, ",")(i)
5                 intCount = intCount + 1
                  
6                 If intCount = 200 Then
7                     stridSQL = stridSQL & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
8                     intCount = 0
9                     strTemp = ""
10                End If
11            Next
12            If strTemp <> "" Then
13                stridSQL = Mid(stridSQL, 12) & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
14            End If
15        Else
16            stridSQL = "Select Column_Value From Table(Cast(F_Num2list([3]) As Zltools.T_Numlist))"
17        End If
          
          '��ͨ
18        strSQL = "Select distinct [�ؼ���] 1 Type,A.����id,a.������Դ ,b.����ʱ��,b.������,b.�����,b.���ʱ��,e.���� ������Ŀ����, " & vbNewLine & _
                  "       D.������ || '(' || D.Ӣ���� || ')' ָ��,c.������, D.��λ," & vbNewLine & _
                  "       Decode(C.�����־, 1, '', 2, '��', 3, '��', 4, '�쳣', 5, '����', 6, '����',7,'��������',8,'��������', '') �����־," & vbNewLine & _
                  "       c.����ο�,C.�������, 0 ��˽��Ŀ,b.�걾����,a.Ӥ�� Ӥ�����,d.ָ�����,a.�걾id,a.���id,d.������,d.Ӣ����,c.�ο���ֵ,c.�ο���ֵ,nvl(d.С��λ��,2) С��λ��,b.������Դ,d.������� " & vbNewLine & _
                  "From ����������� A, ���鱨���¼ B, ���鱨����ϸ C, ����ָ�� D,���������Ŀ e [���]" & vbNewLine & _
                  "Where A.�걾id = B.Id And B.Id = C.�걾id And C.��Ŀid = D.Id And Nvl(B.΢����, 0) <> 1 And" & vbNewLine & _
                  "      a.���id = c.���id and a.���id = e.id  [����1]" & vbNewLine & _
                  "     [����] " & vbNewLine & _
                  " order by a.����id,a.�걾ID,a.���id,c.������� "
19        If intType = 1 Then
20           strSQL = Replace(strSQL, "[����1]", "")
21        Else
22            strSQL = Replace(strSQL, "[����1]", " and   b.����� is not null")
23        End If
24        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[����]", " and A.����id In (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) ")
      '        strSQL = Replace(strSQL, "[����]", " and A.����id In (" & stridSQL & ")")
25            strSQL = Replace(strSQL, "[�ؼ���]", "/*+cardinality(f,10)*/")
26            strSQL = Replace(strSQL, "[���]", ",(" & stridSQL & ") f")
27            strSQL = Replace(strSQL, "[����]", " and a.����id=f.Column_Value")
28        Else
29            strSQL = Replace(strSQL, "[�ؼ���]", "")
30            strSQL = Replace(strSQL, "[���]", "")
31            strSQL = Replace(strSQL, "[����]", " and A.HIS����ID = [1] ")
32        End If

33        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatient, Mid(strTemp, 2), strAdvices)
34        lngID = 0
35        Do Until rsTmp.EOF
              '���ݸ�ʽ���
36            str�ο���ֵ = IIf(IsNumeric(rsTmp("�ο���ֵ") & ""), IIf(Mid(rsTmp("�ο���ֵ") & "", 1, 1) = ".", "0" & rsTmp("�ο���ֵ"), rsTmp("�ο���ֵ") & ""), rsTmp("�ο���ֵ") & "")
37            str�ο���ֵ = IIf(IsNumeric(rsTmp("�ο���ֵ") & ""), IIf(Mid(rsTmp("�ο���ֵ") & "", 1, 1) = ".", "0" & rsTmp("�ο���ֵ"), rsTmp("�ο���ֵ") & ""), rsTmp("�ο���ֵ") & "")
              If IsNumeric(rsTmp("������") & "") Then
38               str������ = IIf(Val(rsTmp("�������") & "") = 1, Format(rsTmp("������") & "", IIf(Val(rsTmp("С��λ��") & "") > 0, "0." & String(rsTmp("С��λ��") & "", "0"), "0")), rsTmp("������") & "")
              Else
                  str������ = rsTmp("������") & ""
              End If
39            If InStr(rsTmp("����ο�") & "", "--") > 0 Then
40                str����ο� = IIf(IsNumeric(Split(rsTmp("����ο�") & "", "--")(0)), IIf(Mid(Split(rsTmp("����ο�") & "", "--")(0), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "--")(0), Split(rsTmp("����ο�") & "", "--")(0)), Split(rsTmp("����ο�") & "", "--")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("����ο�") & "", "--")(1)), IIf(Mid(Split(rsTmp("����ο�") & "", "--")(1), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "--")(1), Split(rsTmp("����ο�") & "", "--")(1)), Split(rsTmp("����ο�") & "", "--")(1))
41            ElseIf InStr(rsTmp("����ο�") & "", "-") > 0 Then
42                str����ο� = IIf(IsNumeric(Split(rsTmp("����ο�") & "", "-")(0)), IIf(Mid(Split(rsTmp("����ο�") & "", "-")(0), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "-")(0), Split(rsTmp("����ο�") & "", "-")(0)), Split(rsTmp("����ο�") & "", "-")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("����ο�") & "", "-")(1)), IIf(Mid(Split(rsTmp("����ο�") & "", "-")(1), 1, 1) = ".", "0" & Split(rsTmp("����ο�") & "", "-")(1), Split(rsTmp("����ο�") & "", "-")(1)), Split(rsTmp("����ο�") & "", "-")(1))
43            Else
44                str����ο� = rsTmp("����ο�") & ""
45            End If
              
46            If lngID <> NVL(rsTmp("����ID"), 0) Or lngSampleID <> NVL(rsTmp("�걾ID"), 0) Or lngSampleGroup <> NVL(rsTmp("���id"), 0) Then
47                strSampleOne = strSampleOne & conSplit1 & "1" & conSplit2 & rsTmp("����ID") & conSplit2 & rsTmp("������Դ") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("����ʱ��")), "", rsTmp("����ʱ��"))) & conSplit2 & _
                              rsTmp("������") & conSplit2 & rsTmp("�����") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("���ʱ��")), "", rsTmp("���ʱ��"))) & conSplit2 & rsTmp("������Ŀ����") & conSplit2 & _
                              rsTmp("�걾����") & conSplit2 & NVL(rsTmp("Ӥ�����"), "0") & conSplit2 & rsTmp("ָ��") & conSplit4 & str������ & conSplit4 & rsTmp("��λ") & _
                              conSplit4 & rsTmp("�����־") & conSplit4 & str����ο� & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("��˽��Ŀ") & conSplit4 & rsTmp("ָ�����") & _
                              conSplit4 & rsTmp("������") & conSplit4 & rsTmp("Ӣ����") & conSplit4 & str�ο���ֵ & conSplit4 & str�ο���ֵ & IIf(rsTmp("������Դ") & "" = "4", conSplit4 & rsTmp("С��λ��"), "")
48            Else
49                strSampleOne = strSampleOne & conSplit3 & rsTmp("ָ��") & conSplit4 & str������ & conSplit4 & rsTmp("��λ") & _
                              conSplit4 & rsTmp("�����־") & conSplit4 & str����ο� & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("��˽��Ŀ") & conSplit4 & rsTmp("ָ�����") & _
                              conSplit4 & rsTmp("������") & conSplit4 & rsTmp("Ӣ����") & conSplit4 & str�ο���ֵ & conSplit4 & str�ο���ֵ & IIf(rsTmp("������Դ") & "" = "4", conSplit4 & rsTmp("С��λ��"), "")
50            End If
51            lngID = NVL(rsTmp("����ID"), 0)
52            lngSampleID = NVL(rsTmp("�걾ID"), 0)
53            lngSampleGroup = NVL(rsTmp("���id"), 0)
54            rsTmp.MoveNext
55        Loop
          
          
56        lngID = 0
57        lngMicroID = 0
58        strSQL = "Select distinct [�ؼ���] 2 Type,A.����id,a.������Դ ,b.����ʱ��,b.������,b.�����,b.���ʱ��,g.���� ������Ŀ����, " & vbNewLine & _
                  "       E.������ || '(' || E.Ӣ���� || ')' ϸ��, C.��������, C.��ҩ����, F.������ || '(' || F.Ӣ���� || ')' ������, D.��� �����ؽ��," & vbNewLine & _
                  "       D.�������, D.ҩ������, F.�÷�����1, F.�÷�����2, ѪҩŨ��1, ѪҩŨ��2, ��ҩŨ��1, ��ҩŨ��2,c.ϸ��ID,b.�걾����,a.Ӥ�� Ӥ�����" & vbNewLine & _
                  "From ����������� A, ���鱨���¼ B, ���鱨��ϸ�� C, ���鱨��ҩ�� D, ����ϸ����¼ E, ����ҩ�� F,���������Ŀ g [���]" & vbNewLine & _
                  "Where A.�걾id = B.Id And B.΢���� = 1 And B.Id = C.�걾id And C.Id = D.���id And C.ϸ��id = E.Id And D.ҩ��id = F.Id and a.���id = g.id " & vbNewLine & _
                  "      [����] " & vbNewLine & _
                  " order by a.����id,c.ϸ��id"
                  
59        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[����]", " and A.����id In (" & stridSQL & ")")
60            strSQL = Replace(strSQL, "[�ؼ���]", "/*+cardinality(h,10)*/")
61            strSQL = Replace(strSQL, "[���]", ",(" & stridSQL & ") h")
62            strSQL = Replace(strSQL, "[����]", " and a.����id=h.Column_Value")
63        Else
64            strSQL = Replace(strSQL, "[�ؼ���]", "")
65            strSQL = Replace(strSQL, "[���]", "")
66            strSQL = Replace(strSQL, "[����]", " and A.HIS����ID = [1] ")
67        End If
68        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatient, Mid(strTemp, 2), strAdvices)
          
          
          '               ����(2=΢����)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>Ӥ�����<split2>
          '               ϸ����1<split3>����1<split3>��ҩ����1<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split2>
          '               ϸ����2<split3>����2<split3>��ҩ����2<split3>
          '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
          '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split1>
          
          
69        If rsTmp.RecordCount <= 0 Then
70            strSQL = "Select  distinct [�ؼ���] 2 Type, a.����id, a.������Դ, b.����ʱ��, b.������, b.�����, b.���ʱ��, c.������ ������Ŀ����, c.δ��� ϸ��, c.��������, c.��ҩ����, '' ������, '' �����ؽ��, '' �������," & vbNewLine & _
                       "          '' ҩ������, '' �÷�����1, '' �÷�����2, '' ѪҩŨ��1, '' ѪҩŨ��2, '' ��ҩŨ��1, '' ��ҩŨ��2, c.ϸ��id, b.�걾����, a.Ӥ�� Ӥ�����" & vbNewLine & _
                       "   From ����������� A, ���鱨���¼ B, ���鱨��ϸ�� C [���]" & vbNewLine & _
                       "   Where a.�걾id = b.Id And b.΢���� = 1 And b.Id = c.�걾id " & vbNewLine & _
                       "      [����] " & vbNewLine & _
                       "   Order By a.����id, c.ϸ��id"
                      
71            If lngPatient = 0 Then
      '            strSQL = Replace(strSQL, "[����]", " and A.����id In (" & stridSQL & ")")
72                strSQL = Replace(strSQL, "[�ؼ���]", "/*+cardinality(d,10)*/")
73                strSQL = Replace(strSQL, "[���]", ",(" & stridSQL & ") D")
74                strSQL = Replace(strSQL, "[����]", " and a.����id=d.Column_Value")
75            Else
76                strSQL = Replace(strSQL, "[�ؼ���]", "")
77                strSQL = Replace(strSQL, "[���]", "")
78                strSQL = Replace(strSQL, "[����]", " and A.HIS����ID = [1] ")
79            End If
80            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatient, Mid(strTemp, 2), strAdvices)
81        End If
          
82        Do Until rsTmp.EOF
83            If lngID <> NVL(rsTmp("����ID"), 0) Then
84                strSampleTwo = strSampleTwo & conSplit1 & "2" & conSplit2 & IIf(IsNull(rsTmp("����ID")), "", rsTmp("����ID")) & conSplit2 & IIf(IsNull(rsTmp("������Դ")), "", rsTmp("������Դ")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("����ʱ��")), "", rsTmp("����ʱ��"))) & conSplit2 & _
                              IIf(IsNull(rsTmp("������")), "", rsTmp("������")) & conSplit2 & IIf(IsNull(rsTmp("�����")), "", rsTmp("�����")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("���ʱ��")), "", rsTmp("���ʱ��"))) & conSplit2 & IIf(IsNull(rsTmp("������Ŀ����")), "", rsTmp("������Ŀ����")) & conSplit2 & _
                              IIf(IsNull(rsTmp("�걾����")), "", rsTmp("�걾����")) & conSplit2 & NVL(rsTmp("Ӥ�����"), "0") & conSplit2 & IIf(IsNull(rsTmp("ϸ��")), "", rsTmp("ϸ��")) & conSplit3 & IIf(IsNull(rsTmp("��������")), "", rsTmp("��������")) & conSplit3 & IIf(IsNull(rsTmp("��ҩ����")), "", rsTmp("��ҩ����")) & _
                              conSplit3 & IIf(IsNull(rsTmp("������")), "", rsTmp("������")) & conSplit4 & IIf(IsNull(rsTmp("�����ؽ��")), "", rsTmp("�����ؽ��")) & conSplit4 & IIf(IsNull(rsTmp("�������")), "", rsTmp("�������")) & conSplit4 & IIf(IsNull(rsTmp("ҩ������")), "", rsTmp("ҩ������")) & conSplit4 & IIf(IsNull(rsTmp("�÷�����1")), "", rsTmp("�÷�����1")) & _
                              conSplit4 & IIf(IsNull(rsTmp("�÷�����2")), "", rsTmp("�÷�����2")) & conSplit4 & IIf(IsNull(rsTmp("ѪҩŨ��1")), "", rsTmp("ѪҩŨ��1")) & conSplit4 & IIf(IsNull(rsTmp("ѪҩŨ��2")), "", rsTmp("ѪҩŨ��2")) & conSplit4 & IIf(IsNull(rsTmp("��ҩŨ��1")), "", rsTmp("��ҩŨ��1")) & conSplit4 & IIf(IsNull(rsTmp("��ҩŨ��2")), "", rsTmp("��ҩŨ��2"))
85                lngMicroID = NVL(rsTmp("ϸ��ID"), 0)
86            Else
87                If lngMicroID <> NVL(rsTmp("ϸ��ID"), 0) Then
88                    strSampleTwo = strSampleTwo & conSplit2 & rsTmp("ϸ��") & conSplit3 & rsTmp("��������") & conSplit3 & rsTmp("��ҩ����") & _
                              conSplit3 & rsTmp("������") & conSplit4 & rsTmp("�����ؽ��") & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("ҩ������") & conSplit4 & rsTmp("�÷�����1") & _
                              conSplit4 & rsTmp("�÷�����2") & conSplit4 & rsTmp("ѪҩŨ��1") & conSplit4 & rsTmp("ѪҩŨ��2") & conSplit4 & rsTmp("��ҩŨ��1") & conSplit4 & rsTmp("��ҩŨ��2")
89                Else
90                    strSampleTwo = strSampleTwo & conSplit3 & rsTmp("������") & conSplit4 & rsTmp("�����ؽ��") & conSplit4 & rsTmp("�������") & conSplit4 & rsTmp("ҩ������") & _
                              conSplit4 & rsTmp("�÷�����1") & conSplit4 & rsTmp("�÷�����2") & conSplit4 & rsTmp("ѪҩŨ��1") & conSplit4 & rsTmp("ѪҩŨ��2") & _
                              conSplit4 & rsTmp("��ҩŨ��1") & conSplit4 & rsTmp("��ҩŨ��2")
91                End If
92            End If
93            lngID = NVL(rsTmp("����ID"), 0)
94            lngMicroID = NVL(rsTmp("ϸ��ID"), 0)
95            rsTmp.MoveNext
96        Loop

97        If strSampleOne <> "" Then
98            strSampleOne = Mid(strSampleOne, Len(conSplit1) + 1)
99        End If
100       If strSampleTwo <> "" Then
101           strSampleTwo = Mid(strSampleTwo, Len(conSplit1) + 1)
102       End If
103       If strSampleTwo <> "" Then
104           If strSampleOne = "" Then
105               CreateLisValueStrForTJ = strSampleTwo
106           Else
107               CreateLisValueStrForTJ = strSampleOne & conSplit1 & strSampleTwo
108           End If
109       Else
110           CreateLisValueStrForTJ = strSampleOne
111       End If


112       Exit Function
CreateLisValueStrForTJ_Error:
113       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(CreateLisValueStr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
114       Err.Clear
          
End Function

Public Function DelLisApplication(strAdvices As String, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                   ɾ��LIS�ж�Ӧ���뵥
          '����                   strAdvices ҽ������,��ʽ
          '                       ��ʽ��<�ɼ�ҽ��1,�ɼ�ҽ��2,.....>
          '                       strErr �д�����Ϣʱ���ش�����Ϣ
          '����                   TRUE=���ͳɹ� FALSE=����ʧ��
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim strSQL As String
              
1         On Error GoTo DelLisApplication_Error

2         strSQL = "Zl_�������뵥_Delete('" & strAdvices & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "ɾ�����뵥")
4         DelLisApplication = True

5         Exit Function
DelLisApplication_Error:
6         strErr = "ɾ�����뵥����" & Err.Number & " " & Err.Description
7         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(DelLisApplication)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
8         Err.Clear
End Function


Public Function SendLisApplication(strAdvices As String, strDiagnose As String, Optional strErr As String) As Boolean
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����                   ���ͼ������뵥��LISϵͳ��
      '����                   strAdvices ҽ������,��ʽ
      '                       ��ʽ��<����ҽ��1,�ɼ�ҽ��1,ִ�п��ұ���1,�걾1;����ҽ��2,�ɼ�ҽ��2,ִ�п��ұ���2,�걾2;.....>
      '                       strDiagnose �����Ϣ
      '����                   TRUE=���ͳɹ� FALSE=����ʧ��
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim intloop As Integer
          Dim astrList() As String
          Dim astrItem() As String
          Dim strData As String
          Dim astrSQL() As String
          Dim blnRollBack As Boolean

1         On Error GoTo SendLisApplication_Error

2         astrList = Split(strAdvices, ";")
3         For intloop = 0 To UBound(astrList)
4             astrItem = Split(astrList(intloop), ",")
5             strSQL = "Select  distinct a.Id,a.���id,a.�걾��λ,a.ִ�п���id From ����ҽ����¼ a,����ҽ������ b  Where ���id=[1]  and a.id= b.ҽ��id  and  b.ִ��״̬=0 "
6             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ҽ����¼", astrItem(1))
7             If rsTmp.RecordCount > 0 Then
8                 Do While Not rsTmp.EOF
9                     strData = rsTmp!ID & "," & rsTmp!���ID & "," & rsTmp!ִ�п���id & "," & rsTmp!�걾��λ
10                    blnRollBack = SendLisApplicationAll(strData, strDiagnose, strErr)
11                    If blnRollBack = False Then
12                        SendLisApplication = False
13                        Exit Function   'ҽ������ʧ��SendLisApplicationҲ�����true�����ٴ��Ǳ�û����ʾ
14                    Else
15                        SendLisApplication = True
16                    End If
17                    rsTmp.MoveNext
18                Loop
19            End If
20        Next
21        SendLisApplication = True

22        Exit Function
SendLisApplication_Error:
23        SendLisApplication = False
24        strErr = "����ţ�" & Err.Number & "    ����������" & Err.Description
25        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(SendLisApplication)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
26        Err.Clear
End Function

Public Function SendLisApplicationAll(strAdvices As String, strDiagnose As String, Optional strErr As String) As Boolean
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����                   ���ͼ������뵥��LISϵͳ��
      '����                   strAdvices ҽ������,��ʽ
      '                       ��ʽ��<����ҽ��1,�ɼ�ҽ��1,ִ�п��ұ���1,�걾1;����ҽ��2,�ɼ�ҽ��2,ִ�п��ұ���2,�걾2;.....>
      '                       strDiagnose �����Ϣ
      '����                   TRUE=���ͳɹ� FALSE=����ʧ��
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim rsTmp As ADODB.Recordset
          Dim rsGather As ADODB.Recordset
          Dim rsExpenses As ADODB.Recordset
          Dim rsGatherExpenses As ADODB.Recordset       '�ɼ����ü�¼
          Dim rsBabyName As ADODB.Recordset
          Dim rsItem As ADODB.Recordset
          Dim rsExpensesAdvice As ADODB.Recordset
          Dim rsAdviceAddition As ADODB.Recordset        'ҽ������
          Dim strExpensesAdvice As String
          Dim strAdviceAddition As String                'ҽ���������������
          Dim strBabyName As String
          Dim strBabySex As String
          Dim strBirthDay As String
          Dim strSQL As String
          Dim intloop As Integer
          Dim astrList() As String
          Dim astrItem() As String
          Dim intItem As Integer
          Dim strData As String
          Dim astrSQL() As String
          Dim blnRollBack As Boolean
          Dim strPayState As String
          Dim strPayStateOne As String
          Dim intTypeFind As Integer
          Dim strExpensesItem As String   '������Ŀҽ��
          Dim strRef As String            '�ο�����
          Dim strRefComItem As String
          Dim blnHave As Boolean
          Dim lngRefItemID As Long      '�ο�Ҫ��ID
          'Dim strDiagnose As String


1         On Error GoTo SendLisApplicationAll_Error

          Dim strWriteItem As String
2         strWriteItem = "����id,ҽ��id,������Դ,����id,Ӥ��,����,�Ա�,����,���Ʊ���,�걾����,������,����ʱ��,�������,����,������,���˿���,������־,�Һŵ�,�����,סԺ��,��������,��ҳid," & _
                         "ǩ����,ǩ��ʱ��,������,����ʱ��,�ͼ���,�ͼ�ʱ��,�Ʒ�״̬,����,��������,������ұ���,���˿��ұ���,��������,·��״̬,��������"

3         ReDim Preserve astrSQL(0)
4         astrList = Split(strAdvices, ";")
5         For intloop = 0 To UBound(astrList)
6             astrItem = Split(astrList(intloop), ",")
7             strData = ""
              '�Ȳ��ҵ�ҽ����ص���Ϣ
8             strSQL = "Select A.Id ҽ��id, A.���id ����id, A.����ʱ�� ����ʱ��, a.������Դ, A.����id, A.Ӥ��, C.����, decode(C.�Ա�,'��',1,'Ů',2,'δ֪',9,0) �Ա�, a.����, A.����ҽ�� ������, A.����ʱ�� ����ʱ��, D.���� �������," & vbNewLine & _
                       "       C.��ǰ���� ����, C.������, E.���� ���˿���, A.������־, A.�Һŵ�, C.�����, C.סԺ��, C.��������, A.��ҳid, B.������ ǩ����, B.����ʱ�� ǩ��ʱ��, B.������, B.����ʱ��," & vbNewLine & _
                       "       b.��������,decode(a.������Դ,2,s.��������,c.��������) ��������, decode(a.������Դ,2,s.·��״̬,null)·��״̬,  B.�ͼ���, B.�걾�ͳ�ʱ�� �ͼ�ʱ��,b.�Ʒ�״̬,a.������ĿID,f.���� ���Ʊ���,a.�걾��λ �걾����," & vbNewLine & _
                       "       b.��¼����,e.���� ���˿��ұ���,d.���� ������ұ���,g.���� ��������,g.���� ����,a.Ƥ�Խ�� ���ܷ���ID,a.������� �������� " & vbNewLine & _
                       "From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���ű� D, ���ű� E,������ĿĿ¼ F,���ű� g,������ҳ s " & vbNewLine & _
                       "Where A.Id = B.ҽ��id And A.����id = C.����id And A.��������id = D.Id And A.���˿���id = E.Id and a.������Ŀid = f.id and a.����id= s.����id(+) and a.��ҳid = s.��ҳid(+)" & vbNewLine & _
                       "   and c.��ǰ����ID = g.id(+) and a.id = [1] "

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ҽ��", astrItem(0))

10            If rsTmp.RecordCount > 0 Then
11                If Val(rsTmp("Ӥ��") & "") > 0 Then
12                    strSQL = "Select b.Ӥ������, Decode(Substr(b.Ӥ���Ա�, Instr(b.Ӥ���Ա�, '-')+1),'��',1,'Ů',2,'δ֪',9,0) �Ա�,b.����ʱ��" & vbNewLine & _
                               "   From ����ҽ����¼ A, ������������¼ B" & vbNewLine & _
                               "   Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Ӥ�� = b.��� And" & vbNewLine & _
                               "         a.���id In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And Rownum = 1"
13                    Set rsBabyName = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�ҽ��", astrItem(1))
14                End If
15            End If
16            If Not rsBabyName Is Nothing Then
17                If rsBabyName.RecordCount > 0 Then
18                    strBirthDay = StringFormatDate(rsBabyName("����ʱ��") & "")
19                Else
20                    strBirthDay = StringFormatDate(rsTmp("��������") & "")
21                End If
22            Else
23                strBirthDay = StringFormatDate(rsTmp("��������") & "")
24            End If

              'ҽ������
25            strSQL = "Select b.Ҫ��ID, b.��Ŀ, b.����, b.����" & vbNewLine & _
                       "From ����ҽ������ b" & vbNewLine & _
                       "Where b.ҽ��id In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) " & vbNewLine & _
                       "Order By ҽ��id, ����"
26            Set rsAdviceAddition = ComOpenSQL(Sel_His_DB, strSQL, "ҽ������", astrItem(1))

              'ȡ�ɼ���ʽ�Ͳɼ�����
27            strSQL = "Select B.���� �ɼ���ʽ, C.���� �ɼ����� " & vbNewLine & _
                       "From ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C " & vbNewLine & _
                       "Where A.������Ŀid = B.Id And A.ִ�п���id = C.Id and a.id = [1] "
28            Set rsGather = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ���ʽ", astrItem(1))


29            strSQL = "select ID,������� from ����ҽ����¼ where id in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) " & _
                       " union all " & _
                       "select ID,������� from ����ҽ����¼ where ���id in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) "
30            Set rsExpensesAdvice = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ���ʽ", astrItem(1))
31            strExpensesAdvice = ""
32            strExpensesItem = ""
33            Do Until rsExpensesAdvice.EOF
34                If rsExpensesAdvice("�������") = "E" Then
35                    If strExpensesAdvice = "" Then
36                        strExpensesAdvice = rsExpensesAdvice("ID")
37                    Else
38                        strExpensesAdvice = strExpensesAdvice & "," & rsExpensesAdvice("ID")
39                    End If
40                Else
41                    If strExpensesItem = "" Then
42                        strExpensesItem = rsExpensesAdvice("ID")
43                    Else
44                        strExpensesItem = strExpensesItem & "," & rsExpensesAdvice("ID")
45                    End If
46                End If
47                rsExpensesAdvice.MoveNext
48            Loop
49            intTypeFind = GetAdviceFeeKind(Val(astrItem(1)))
              '�޳��ɼ�����
50            strSQL = "Select �ѱ� ,Sum(Ӧ�ս��) As Ӧ�ս��, Sum(Decode(��¼״̬, 1, ʵ�ս��, 0)) ʵ�ս�� From סԺ���ü�¼ Where ִ��״̬ <> 9   " & vbNewLine & _
                       " and ҽ����� in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by �ѱ� "
51            If intTypeFind = 1 Then
52                strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
53            End If
54            Set rsExpenses = ComOpenSQL(Sel_His_DB, strSQL, "���ҷ���", strExpensesItem)
              '�ɼ�����
55            strSQL = "Select �ѱ�,Sum(Ӧ�ս��) As �ɼ�Ӧ�ս��, Sum(Decode(��¼״̬, 1, ʵ�ս��, 0)) �ɼ�ʵ�ս�� From סԺ���ü�¼ Where ִ��״̬ <> 9   " & vbNewLine & _
                       " and ҽ����� in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by �ѱ� "
56            If intTypeFind = 1 Then
57                strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
58            End If
59            Set rsGatherExpenses = ComOpenSQL(Sel_His_DB, strSQL, "���ҷ���", strExpensesAdvice)
60            strPayState = funFindAdvicePay(astrItem(0), rsTmp("������Դ") & "")
61            If strPayState <> "" Then
62                If InStr(strPayState, "|") > 0 Then
63                    strPayState = 0
64                Else
65                    strPayStateOne = Split(strPayState, ",")(1)
66                End If
67            End If
              'д��������֯
68            astrItem = Split(strWriteItem, ",")

69            For intItem = 0 To UBound(astrItem)
70                If astrItem(intItem) = "����" Then
71                    strData = strData & ",'" & rsTmp(astrItem(intItem)) & "','" & GetAgeMid(0, rsTmp(astrItem(intItem)) & "") & "','" & GetAgeMid(1, rsTmp(astrItem(intItem)) & "") & "'"
72                ElseIf astrItem(intItem) Like "*ʱ��*" Or astrItem(intItem) Like "*����*" Then
73                    If astrItem(intItem) = "��������" Then
74                        If strBirthDay = "" Then
75                            strData = strData & ",null"
76                        Else
77                            strData = strData & "," & strBirthDay
78                        End If
79                    Else
80                        If rsTmp(astrItem(intItem)) & "" = "" Then
81                            strData = strData & ",null"
82                        Else
83                            strData = strData & "," & StringFormatDate(rsTmp(astrItem(intItem)))
84                        End If
85                    End If
86                ElseIf astrItem(intItem) = "�Ʒ�״̬" Then
87                    strData = strData & "," & Val(strPayStateOne)
88                ElseIf astrItem(intItem) = "���Ʊ���" Then
89                    If gUserInfo.NodeNo <> "-" Then
90                        strSQL = "select id from ���������Ŀ where ���Ʊ��� = [1] and (վ��=[2] or վ�� is null)"
91                    Else
92                        strSQL = "select id from ���������Ŀ where ���Ʊ��� = [1] "
93                    End If
94                    Set rsItem = ComOpenSQL(Sel_Lis_DB, strSQL, "ȡ�����Ŀ", rsTmp(astrItem(intItem)) & "", gUserInfo.NodeNo)
95                    If rsItem.RecordCount > 0 Then
96                        strData = strData & ",'" & rsItem("ID") & "'"
97                    Else
98                        strData = strData & ",null"
99                    End If
100               Else
101                   strData = strData & ",'" & rsTmp(astrItem(intItem)) & "'"
102               End If
103           Next
104           If rsExpenses.RecordCount > 0 Then
105               strData = strData & ",'" & rsGather("�ɼ���ʽ") & "','" & rsGather("�ɼ�����") & "','" & rsExpenses("Ӧ�ս��") & "','" & rsExpenses("ʵ�ս��") & _
                            "','" & rsExpenses("�ѱ�") & "'"
106           Else
107               strData = strData & ",'" & rsGather("�ɼ���ʽ") & "','" & rsGather("�ɼ�����") & "','" & "" & "','" & "" & _
                            "','" & "" & "'"
108           End If

              'strDiagnose = GetPatiDiagnose(Val(rsTmp("����id") & ""), Val(rsTmp("��ҳid") & ""), Val(rsTmp("������Դ") & ""))

              '���
109           strData = strData & ",'" & strDiagnose & "'"

110           If Not rsBabyName Is Nothing Then
111               If rsBabyName.RecordCount > 0 Then
112                   strBabyName = rsBabyName("Ӥ������") & ""
113                   strBabySex = rsBabyName("�Ա�") & ""
115               End If
116           End If
117           strData = strData & ",'" & strBabyName & "','" & strBabySex & "'"
118           strAdviceAddition = ""

              '��ҽ�������л�ȡ�ο�����
119           If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then

120               If rsAdviceAddition.RecordCount > 0 Then
121                   Do Until rsAdviceAddition.EOF
122                       strAdviceAddition = strAdviceAddition & rsAdviceAddition("��Ŀ") & ":" & rsAdviceAddition("����") & vbNewLine
123                       If rsAdviceAddition("Ҫ��ID") & "" <> "" And rsAdviceAddition("��Ŀ") & "" <> "" Then
124                           strRefComItem = GetRefItem(Val(rsAdviceAddition("Ҫ��ID") & ""), rsAdviceAddition("��Ŀ") & "", rsAdviceAddition("����") & "", blnHave, lngRefItemID)
125                           If blnHave Then
126                               strRef = strRef & "<Split D>" & lngRefItemID & "<Split E>" & strRefComItem
127                           End If
128                       End If
129                       rsAdviceAddition.MoveNext
130                   Loop
131               End If
132           Else
133               If rsAdviceAddition.RecordCount > 0 Then
134                   Do Until rsAdviceAddition.EOF
135                       strAdviceAddition = strAdviceAddition & rsAdviceAddition("��Ŀ") & ":" & rsAdviceAddition("����") & vbNewLine
136                       rsAdviceAddition.MoveNext
137                   Loop
138               End If
139           End If
140           strData = strData & ",'" & Replace(strAdviceAddition, vbCrLf, "") & "'"
141           If rsGatherExpenses.RecordCount > 0 Then
142               strData = strData & ",'" & rsGatherExpenses("�ɼ�Ӧ�ս��") & "','" & rsGatherExpenses("�ɼ�ʵ�ս��") & "'"
143           Else
144               strData = strData & ",'',''"
145           End If

              '��������
146           If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
147               If rsTmp("���ܷ���ID") & "" <> "" Then
148                   strData = strData & "," & rsTmp("���ܷ���ID") & ",'" & rsTmp("��������") & "'"
149               Else
150                   strData = strData & ",null,null"
151               End If
152           End If

              '�ο�����
153           If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
154               If strRef <> "" Then
155                   If Mid(strRef, 1, Len("<Split D>")) = "<Split D>" Then strRef = Mid(strRef, Len("<Split D>") + 1)
156                   strData = strData & ",'" & strRef & "'"
157               End If
158           End If

159           strData = Mid(strData, 2)
160           If intloop > 0 Then
161               ReDim Preserve astrSQL(UBound(astrSQL) + 1)
162           End If
163           astrSQL(UBound(astrSQL)) = "Zl_�������뵥_Insert(" & strData & ")"
164       Next
165       blnRollBack = True
          '    gcnHisOracle.BeginTrans
166       For intloop = 0 To UBound(astrSQL)
167           Call ComExecuteProc(Sel_Lis_DB, astrSQL(intloop), "����ҽ��������")
168       Next
          '    gcnHisOracle.CommitTrans
169       blnRollBack = False
170       SendLisApplicationAll = True

          '����ˢ�¿��ڸſ�δ������ǩ����
171       Call SendMessage("RefreshDeptSurvey0")

172       Exit Function
SendLisApplicationAll_Error:
173       SendLisApplicationAll = False
174       strErr = "����ţ�" & Err.Number & "    ����������" & Err.Description
175       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(SendLisApplicationAll)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
176       Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-09-20
'��    ��:  Ҫ����ʾ����ת��������
'��    ��:
'           lngID           Ҫ��ID
'           strName         Ҫ����
'           strShow         Ҫ����ʾ����
'��    ��:
'           strShow         �Ƿ����Ҫ��
'           lngRefItemID    Ҫ��ID
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Function GetRefItem(ByVal lngID As Long, ByVal strName As String, ByVal strShow As String, ByRef blnHave As Boolean, ByRef lngRefItemID As Long) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsVal As ADODB.Recordset
          Dim strSQLValue As String
          Dim strArr() As String
          Dim i As Integer

1         On Error GoTo GetRefItem_Error

2         strSQL = "select ֵ��,ֵ����Դ,ID from ����ָ��ο�Ҫ�� where ID=[1] and Ҫ����=[2]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�ο�Ҫ��", lngID, strName)
4         If rsTmp.EOF Then
5             strSQL = "select ֵ��,ֵ����Դ,ID from ����ָ��ο�Ҫ�� where Ҫ����=[1]"
6             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�ο�Ҫ��", strName)
7         End If
          
8         If Not rsTmp.EOF Then
9             blnHave = True
10            lngRefItemID = Val(rsTmp("ID") & "")
11            If Val(rsTmp("ֵ����Դ") & "") > 0 Then
                  'ֵ��ΪSQL
12                strSQLValue = "select * from (" & rsTmp("ֵ��") & ") where ��ʾ����=[1]"
13                Set rsVal = ComOpenSQL(Sel_Lis_DB, strSQLValue, "ֵ��", strShow)
14                If Not rsVal.EOF Then
15                    GetRefItem = rsVal("��������") & ""
16                End If
17            Else
                  'ֵ��Ϊ�ֶ�����
18                strSQLValue = rsTmp("ֵ��") & ""
19                strArr = Split(strSQLValue, "<SP1>")
20                For i = 0 To UBound(strArr)
21                    If Trim(strShow) = Trim(Split(strArr(i), "<SP2>")(0)) Then
22                        GetRefItem = Trim(Split(strArr(i), "<SP2>")(1))
23                        Exit Function
24                    End If
25                Next
26            End If
27        Else
28            blnHave = False
29        End If


30        Exit Function
GetRefItem_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetRefItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
32        Err.Clear
End Function

Private Function GetAgeMid(intType As Integer, strAge As String) As String
    '����           ת������
    '����           0=ȡ�������� 1=ȡ���䵥λ
    
    If intType = 0 Then
        GetAgeMid = Val(strAge)
    Else
        'GetAgeMid = Replace(strAge, Val(strAge), "")
        GetAgeMid = Mid(strAge, Len("" & Val(strAge)) + 1)
    End If
End Function

Public Function GetSampleDeptRS(Optional strErr As String) As ADODB.Recordset
          '����       ȡ�òɼ����ҵ����ݼ�
          '����       �ҵ��Ĳɼ��������ݼ�

          Dim strSQL As String
1         On Error GoTo GetSampleDeptRS_Error

2         strSQL = "Select Distinct C.Id, C.����, C.����" & vbNewLine & _
                      "From ������ĿĿ¼ A, ����ִ�п��� B, ���ű� C" & vbNewLine & _
                      "Where A.��� = 'E' And A.�������� = '6' And A.Id = B.������Ŀid And B.ִ�п���id = C.Id"
3         Set GetSampleDeptRS = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�����")


4         Exit Function
GetSampleDeptRS_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetSampleDeptRS)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
6         Err.Clear

End Function

Public Function GetSampleTypeRS(Optional strErr As String) As ADODB.Recordset
          '����       ȡ�òɼ���Ŀ�����ݼ�
          '����       �ҵ��Ĳɼ���Ŀ���ݼ�

          Dim strSQL As String
1         On Error GoTo GetSampleTypeRS_Error

2         strSQL = "select id,����,���� from ������ĿĿ¼ where ��� = 'E' and �������� = '6' "
          
3         Set GetSampleTypeRS = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�����")


4         Exit Function
GetSampleTypeRS_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetSampleTypeRS)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
6         Err.Clear

End Function

Public Function Get����ִ�п���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng��Ŀid As Long, ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    Getlng����ִ�п���ID As Long, Getstr����ִ�п����� As String, _
    Optional ByVal int��Χ As Integer = 2, Optional strErr As String) As Boolean
          
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim intִ�п��� As Integer
          Dim lng����Ա����ID As Long
          Dim bln�ϰల�� As Boolean
          Dim lngID As Long
          Dim bytDay As Byte
          
          
1         On Error GoTo Get����ִ�п���_Error

2         strSQL = "Select A.Id, A.ִ�п��� From ������ĿĿ¼ A, �����÷����� B Where A.Id = B.�÷�id And B.��Ŀid = [1] And a.������� IN([2],3)"

3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", lng��Ŀid, int��Χ)
4         If rsTmp.RecordCount > 0 Then
5             lngID = rsTmp("ID")
6             intִ�п��� = rsTmp("ִ�п���")
7         End If
          
8         Select Case intִ�п���
              Case 0, 5 '0-��ִ�еĶ���,5-Ժ��ִ��
9                 Get����ִ�п��� = True: Exit Function
10            Case 1 '1-�������ڿ���
11                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([1]) Order by ����"
12            Case 2 '2-�������ڲ���
13                If int��Χ = 1 Then
14                    strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([1]) Order by ����"
15                Else
16                    strSQL = _
                          " Select A.ID,A.����,A.����,A.����" & _
                          " From ���ű� A,������ҳ B" & _
                          " Where A.ID=B.��ǰ����ID And B.����ID=[2] And B.��ҳID=[3] "
17                End If
18            Case 3 '3-����Ա���ڿ���
19                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([4]) Order by ����"
20                lng����Ա����ID = Get����Ա����ID(int��Χ)
21            Case 4 '4-ָ������
22                bln�ϰల�� = Check�ϰల��(False)
23                If Not bln�ϰల�� Then
24                    strSQL = _
                          " Select Distinct A.ID,A.����,A.����,A.����" & _
                          " From ���ű� A,����ִ�п��� B,��������˵�� C" & _
                          " Where A.ID=B.ִ�п���ID And B.������ĿID=[5] And A.ID=C.����ID" & _
                          " And C.������� IN([6],3) And (B.������Դ is NULL Or B.������Դ=[6])" & _
                          " And (B.��������ID is NULL Or B.��������ID=[1])" & _
                          " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)"
25                Else
26                    bytDay = Weekday(Currentdate, vbMonday) Mod 7 '0=����,1=��һ
27                    strSQL = _
                          " Select Distinct C.ID,C.����,C.����,C.����" & _
                          " From ����ִ�п��� A,���Ű��� B,���ű� C,��������˵�� D" & _
                          " Where A.ִ�п���ID+0=B.����ID And B.����ID=C.ID " & _
                          " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                          " And C.ID=D.����ID And D.������� IN([6],3) And (A.������Դ is NULL Or A.������Դ=[6])" & _
                          " And (A.��������ID is NULL Or A.��������ID=[1]) And A.������ĿID=[5]" & _
                          " And (C.վ��='" & gUserInfo.NodeNo & "' Or C.վ�� is Null)"
28                End If
29            Case 6 '6-���������ڿ���
30                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([11],[6]) Order by ����"
31        End Select
32        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlCISKernel", lng���˿���ID, lng����ID, lng��ҳID, lng����Ա����ID, lngID, int��Χ)
33        If Not rsTmp.EOF Then
34            Getlng����ִ�п���ID = rsTmp!ID
35            Getstr����ִ�п����� = rsTmp!����
36            rsTmp.Filter = "ID=" & lng���˿���ID
37            If rsTmp.EOF Then rsTmp.Filter = "ID=" & lng���˿���ID
      '        If rsTmp.EOF And int��Χ = 2 Then rsTmp.Filter = "ID=" & Get����ID(lng����ID, lng��ҳId)
38            If Not rsTmp.EOF Then Getlng����ִ�п���ID = rsTmp!ID: Getstr����ִ�п����� = rsTmp!����
39        End If
40        rsTmp.Filter = ""
          'Set Get����ִ�п��� = rsTmp


41        Exit Function
Get����ִ�п���_Error:
42        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(Get����ִ�п���)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
43        Err.Clear

End Function

Public Function Get����ִ�п���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lng��Ŀid As Long, ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    Getlng����ִ�п���ID As Long, Getstr����ִ�п����� As String, _
    Optional ByVal int��Χ As Integer = 2, Optional strErr As String) As Boolean
          
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim intִ�п��� As Integer
          Dim lng����Ա����ID As Long
          Dim bln�ϰల�� As Boolean
          Dim lngID As Long
          Dim bytDay As Byte
          
1         On Error GoTo Get����ִ�п���_Error

2         strSQL = "Select A.Id, A.ִ�п��� From ������ĿĿ¼ A  Where A.Id =   [1]"

3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", lng��Ŀid)
4         If rsTmp.RecordCount > 0 Then
5             lngID = rsTmp("ID")
6             intִ�п��� = rsTmp("ִ�п���")
7         End If
          
8         Select Case intִ�п���
              Case 0, 5 '0-��ִ�еĶ���,5-Ժ��ִ��
9                 Get����ִ�п��� = True: Exit Function
10            Case 1 '1-�������ڿ���
11                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([1]) Order by ����"
12            Case 2 '2-�������ڲ���
13                If int��Χ = 1 Then
14                    strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([1]) Order by ����"
15                Else
16                    strSQL = _
                          " Select A.ID,A.����,A.����,A.����" & _
                          " From ���ű� A,������ҳ B" & _
                          " Where A.ID=B.��ǰ����ID And B.����ID=[2] And B.��ҳID=[3] "
17                End If
18            Case 3 '3-����Ա���ڿ���
19                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([4]) Order by ����"
20                lng����Ա����ID = Get����Ա����ID(int��Χ)
21            Case 4 '4-ָ������
22                bln�ϰల�� = Check�ϰల��(False)
23                If Not bln�ϰల�� Then
24                    strSQL = _
                          " Select Distinct A.ID,A.����,A.����,A.����" & _
                          " From ���ű� A,����ִ�п��� B,��������˵�� C" & _
                          " Where A.ID=B.ִ�п���ID And B.������ĿID=[5] And A.ID=C.����ID" & _
                          " And C.������� IN([6],3) And (B.������Դ is NULL Or B.������Դ=[6])" & _
                          " And (B.��������ID is NULL Or B.��������ID=[1])" & _
                          " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)"
25                Else
26                    bytDay = Weekday(Currentdate, vbMonday) Mod 7 '0=����,1=��һ
27                    strSQL = _
                          " Select Distinct C.ID,C.����,C.����,C.����" & _
                          " From ����ִ�п��� A,���Ű��� B,���ű� C,��������˵�� D" & _
                          " Where A.ִ�п���ID+0=B.����ID And B.����ID=C.ID " & _
                          " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.��ʼʱ��,'HH24:MI:SS') and To_Char(B.��ֹʱ��,'HH24:MI:SS') " & _
                          " And C.ID=D.����ID And D.������� IN([6],3) And (A.������Դ is NULL Or A.������Դ=[6])" & _
                          " And (A.��������ID is NULL Or A.��������ID=[1]) And A.������ĿID=[5]" & _
                          " And (C.վ��='" & gUserInfo.NodeNo & "' Or C.վ�� is Null)"
28                End If
29            Case 6 '6-���������ڿ���
30                strSQL = "Select ID,����,����,���� From ���ű� Where ID IN([11],[6]) Order by ����"
31        End Select
32        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlCISKernel", lng���˿���ID, lng����ID, lng��ҳID, lng����Ա����ID, lngID, int��Χ)
33        If Not rsTmp.EOF Then
34            Getlng����ִ�п���ID = rsTmp!ID
35            Getstr����ִ�п����� = rsTmp!����
      '        rsTmp.Filter = "ID=" & lng���˿���ID
      '        If rsTmp.EOF Then rsTmp.Filter = "ID=" & lng���˿���ID
      '        If rsTmp.EOF And int��Χ = 2 Then rsTmp.Filter = "ID=" & Get����ID(lng����ID, lng��ҳId)
36            If Not rsTmp.EOF Then Getlng����ִ�п���ID = rsTmp!ID: Getstr����ִ�п����� = rsTmp!����
37        End If
38        rsTmp.Filter = ""


39        Exit Function
Get����ִ�п���_Error:
40        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(Get����ִ�п���)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
41        Err.Clear

End Function

Public Function Get����Ա����ID(ByVal int������� As Integer, Optional strErr As String) As Long
      '���ܣ�ȡ����Ա���������ָ������Ĳ��ţ�ȱʡ��������
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, blnNew As Boolean
          
1         On Error GoTo Get����Ա����ID_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
          
7         If blnNew Then
8             strSQL = "Select Distinct B.����ID,Nvl(B.ȱʡ,0) as ȱʡ,C.������� From ������Ա B,��������˵�� C" & _
                  " Where B.��ԱID = [1] And B.����ID=C.����ID" & _
                  " Order by ȱʡ Desc"
9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlLisHisComm", gUserInfo.ID)
10        End If
11        rsTmp.Filter = "������� = 3 or ������� = " & int�������
          
12        If Not rsTmp.EOF Then
13            Get����Ա����ID = rsTmp!����ID
14        Else
15            Get����Ա����ID = gUserInfo.DeptID
16        End If


17        Exit Function
Get����Ա����ID_Error:
18        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(Get����Ա����ID)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
19        Err.Clear

End Function

Public Function Check�ϰల��(ByVal blnҩ�� As Boolean, Optional strErr As String) As Boolean
      '���ܣ����ҽԺ�Ŀ����Ƿ�ʹ�����ϰల��
      '������blnҩ��=�Ǽ��ҩ���ϰ໹����������
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String
          Static blnҩ��Load As Boolean
          Static blnҩ��Last As Boolean
          Static bln��ҩLoad As Boolean
          Static bln��ҩLast As Boolean
          
1         On Error GoTo Check�ϰల��_Error

2         If blnҩ�� Then '�Ƿ��а���ֻ���ȡһ��
3             If blnҩ��Load Then Check�ϰల�� = blnҩ��Last: Exit Function
4         Else
5             If bln��ҩLoad Then Check�ϰల�� = bln��ҩLast: Exit Function
6         End If
              
7         If blnҩ�� Then
8             strSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
                  " Where A.����ID=B.����ID  And Rownum<2"
9         Else
10            strSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
                  " Where A.����ID=B.����ID  And Rownum<2"
11        End If
12        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "Check�ϰల��")
13        Check�ϰల�� = rsTmp.RecordCount > 0
          
14        If blnҩ�� Then
15            blnҩ��Load = True: blnҩ��Last = Check�ϰల��
16        Else
17            bln��ҩLoad = True: bln��ҩLast = Check�ϰల��
18        End If


19        Exit Function
Check�ϰల��_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(Check�ϰల��)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Function

Public Function SampleBarcodeUpdate(strAdvices As String, strBarCode As String, Optional strSamplingName As String, Optional strErr As String, Optional ByVal intContinue As Integer) As Boolean
          '����                   �ڲ���ʱ���ʱд���������뵽�����¼�У�ȡ��д��յ�������Ϣ
          '
          '����                   strAdvices   ҽ����,���ҽ��ʹ��","�ŷָ�
          '                       strBarCode   Ҫд�������
          '                       ��ʽ:��ҽ��1,ҽ��2,ҽ��3,..
          '                       intContinue 1=�ò�����,2=ȡ������,3=�ز�����,4=��������,5=ȡ���ɼ�,6-ȡ������,7-��ɲɼ�,8-ȡ���ɼ�������,9-��ɲɼ�����������
          '����                   �������Falseʱ����strErr����ʾ�������
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim intloop As Integer
          Dim strbuff As String
          Dim rsbuff As ADODB.Recordset, rsTmp As ADODB.Recordset
          Dim varAdvices As Variant
          Dim strAdvice As String
          Dim blnGet As Boolean
          Dim strProgram As String
          
1         On Error GoTo SampleBarcodeUpdate_Error
          
2         If Mid(strAdvices, Len(strAdvices), 1) = "," Then
3             strAdvices = Mid(strAdvices, 1, Len(strAdvices) - 1)
4         End If
          
5         varAdvices = Split(strAdvices, ",")
6         For intloop = 0 To UBound(varAdvices)
7             If Val(varAdvices(intloop)) <> 0 Then
8                 strbuff = "Select id from ����������� where ҽ��id =[1] or ����id =[1]"
9                 Set rsbuff = ComOpenSQL(Sel_Lis_DB, strbuff, "�ɼ�վ��ѯ", varAdvices(intloop))
10                If rsbuff.EOF Then
                      '�����������뵥����
11                    strbuff = "Select  distinct a.Id,a.���id,a.�걾��λ,a.ִ�п���id From ����ҽ����¼ a,����ҽ������ b  Where id=[1]  and a.id= b.ҽ��id  and  b.ִ��״̬=0 "
12                    Set rsTmp = ComOpenSQL(Sel_His_DB, strbuff, "����ҽ����¼", varAdvices(intloop))
13                    If rsTmp.RecordCount > 0 Then
14                        If Val(rsTmp!���ID & "") <> 0 Then
15                            strAdvice = rsTmp!ID & "," & rsTmp!���ID & "," & rsTmp!ִ�п���id & "," & rsTmp!�걾��λ
16                            blnGet = SendLisApplication(strAdvice, "", strErr)
17                        End If
18                    End If
19                End If
20            End If
21        Next
          
22        If VerCompare(gSysInfo.VersionLIS, "10.35.150") <> -1 Then
23            If strSamplingName = "" Then
24                If intContinue = 2 Then
                      '���պ�ȡ�����
25                    strSQL = "Zl_������������_Updatenew(0,'" & strAdvices & "','" & strBarCode & "',Null,2)"
26                ElseIf intContinue = 4 Then
                      '��������
27                    strSQL = "Zl_������������_Updatenew(1,'" & strAdvices & "','" & strBarCode & "')"
28                ElseIf intContinue = 6 Then
                      'ȡ������
29                    strSQL = "Zl_������������_Updatenew(2,'" & strAdvices & "','" & strBarCode & "')"
30                ElseIf intContinue = 5 Then
                      'ȡ���ɼ�
31                    strSQL = "Zl_������������_Updatenew(4,'" & strAdvices & "','" & strBarCode & "')"
32                ElseIf intContinue = 8 Then
                      'ȡ���ɼ�������
33                    strSQL = "Zl_������������_Updatenew(6,'" & strAdvices & "','" & strBarCode & "')"
34                End If
35            Else
36                If intContinue = 7 Then
                      '��ɲɼ�
37                    strSQL = "Zl_������������_Updatenew(3,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "')"
38                ElseIf intContinue = 9 Then
                      '�����������ɲɼ�
39                    strSQL = "Zl_������������_Updatenew(5,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "')"
40                Else
                      '���պ��ò����ز�
41                    strSQL = "Zl_������������_Updatenew(0,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "'," & intContinue & ")"
42                End If
43            End If
44        Else
              '--0=д������ �ɼ��˵ȣ�1=ȡ���ɼ��ˣ�2-����д�����룬3-ȡ�����룬4-��ɲɼ�
45            If strSamplingName = "" Then
46                If intContinue = 5 Then
47                    strSQL = "Zl_������������_Update(5,'" & strAdvices & "','" & strBarCode & "')"
48                ElseIf intContinue = 4 Then
49                    strSQL = "Zl_������������_Update(2,'" & strAdvices & "','" & strBarCode & "')"
50                ElseIf intContinue = 6 Then
51                    strSQL = "Zl_������������_Update(3,'" & strAdvices & "','" & strBarCode & "')"
52                Else
53                     strSQL = "Zl_������������_Update(1,'" & strAdvices & "','" & strBarCode & "')"
54                End If
55            Else
56                If intContinue = 4 Then
57                    strSQL = "Zl_������������_Update(2,'" & strAdvices & "','" & strBarCode & "')"
58                ElseIf intContinue = 7 Then
59                    strSQL = "Zl_������������_Update(4,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "')"
60                Else
61                    strSQL = "Zl_������������_Update(0,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "'," & intContinue & ")"
62                End If
63            End If
64        End If
          
65        Call ComExecuteProc(Sel_Lis_DB, strSQL, "д������")
          
          '����ҵ���µ��ò����顢ȡ����ɡ��ز���Ϣ
66        If intContinue = 1 Or intContinue = 2 Or intContinue = 3 Then
67            If funWriteInLisNotify(1, strAdvices, intContinue, strErr) = False Then Exit Function
68        End If
          
69        SampleBarcodeUpdate = True
          
          '����ˢ�¿��ڸſ�δ�ͼ��ǩ����
70        Call SendMessage("RefreshDeptSurvey1")
          
71        Exit Function
SampleBarcodeUpdate_Error:
72        SampleBarcodeUpdate = False
73        strErr = "д�����뵥�������" & Err.Number & " " & Err.Description
74        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(SampleBarcodeUpdate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
75        Err.Clear
End Function

Public Function funGetItemMoney(strAdvices As String, Optional strErr As String) As Boolean
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����               ͨ��ҽ��ID�õ���ǰ��Ŀ�շѵĽ��
      '����               strAdvices       ���ҽ��ID�ö��ŷָ�
      '                   strErr           ����д�����Ϣʱ���ش�����Ϣ
      '����               �շ�״̬,Ӧ�ս��,ʵ�ս��
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsPatient As ADODB.Recordset
          Dim rsGatherExpenses As ADODB.Recordset
          Dim strItem As String
          Dim intPatient As Integer
          Dim intType As Integer
          Dim astrAdvice() As String
          Dim intloop As Integer
          Dim astrSQL() As String
          Dim intTypeFind As Integer
          Dim strFindAdice As String

1         On Error GoTo funGetItemMoney_Error

2         ReDim Preserve astrSQL(0)

3         astrAdvice = Split(strAdvices, ",")

4         For intloop = 0 To UBound(astrAdvice)
5             strSQL = "select a.������Դ,a.id,b.��¼����,a.������� from ����ҽ����¼ a,����ҽ������ b where a.id = b.ҽ��id and a.���id = [1] "
6             Set rsPatient = ComOpenSQL(Sel_His_DB, strSQL, "���ҷ���", Val(astrAdvice(intloop)))
7             If rsPatient.RecordCount > 0 Then
8                 Do Until rsPatient.EOF
9                     If rsPatient("��¼����") = "E" Then
10                        strItem = strItem & "," & rsPatient("ID")
11                    Else
12                        strFindAdice = strFindAdice & "," & rsPatient("ID")
13                    End If
14                    intPatient = Val(rsPatient("������Դ") & "")
15                    intType = Val(rsPatient("��¼����") & "")
16                    rsPatient.MoveNext
17                Loop
18                intTypeFind = GetAdviceFeeKind(Val(astrAdvice(intloop)))

19                If strItem <> "" Then
20                    strItem = Val(astrAdvice(intloop)) & strItem
21                Else
22                    strItem = Val(astrAdvice(intloop))
23                End If
                  '�޳��ɼ�����
24                strSQL = "Select /*+ rule */ ��¼״̬,Sum(Ӧ�ս��) As Ӧ�ս��, Sum(Decode(��¼״̬, 1, ʵ�ս��, 0)) ʵ�ս�� From סԺ���ü�¼ Where ִ��״̬ <> 9   " & vbNewLine & _
                         " and ҽ����� in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by ��¼״̬ "
25                If intTypeFind = 1 Then
26                    strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
27                End If
28                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���ҷ���", strFindAdice)
                  '�ɼ�����
29                strSQL = "Select /*+ rule */ �շ����,Sum(Ӧ�ս��) As �ɼ�Ӧ�ս��, Sum(Decode(��¼״̬, 1, ʵ�ս��, 0)) �ɼ�ʵ�ս�� From סԺ���ü�¼ Where ִ��״̬ <> 9   " & vbNewLine & _
                         " and ҽ����� in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by �շ���� "
30                If intTypeFind = 1 Then
31                    strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
32                End If
33                Set rsGatherExpenses = ComOpenSQL(Sel_His_DB, strSQL, "���ҷ���", strItem)

34                If rsTmp.RecordCount > 0 Then
35                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
36                    If rsGatherExpenses.RecordCount > 0 Then
37                        astrSQL(UBound(astrSQL)) = "Zl_���뵥���_Update('" & Val(astrAdvice(intloop)) & "','" & rsTmp("Ӧ�ս��") & "','" & _
                                                     rsTmp("ʵ�ս��") & "','" & rsTmp("��¼״̬") & "','" & rsGatherExpenses("�ɼ�Ӧ�ս��") & "','" & rsGatherExpenses("�ɼ�ʵ�ս��") & "')"
38                    Else
39                        astrSQL(UBound(astrSQL)) = "Zl_���뵥���_Update('" & Val(astrAdvice(intloop)) & "','" & rsTmp("Ӧ�ս��") & "','" & _
                                                     rsTmp("ʵ�ս��") & "','" & rsTmp("��¼״̬") & "')"
40                    End If
41                End If
42                strItem = ""
43               strFindAdice = ""
44            End If
45        Next
46        For intloop = 0 To UBound(astrSQL)
47            If astrSQL(intloop) <> "" Then
48                Call ComExecuteProc(Sel_Lis_DB, astrSQL(intloop), "���·�����Ϣ")
49            End If
50        Next
          
51        funGetItemMoney = True

52        Exit Function
funGetItemMoney_Error:
53        strErr = "����ҽ�����ô���" & Err.Number & " " & Err.Description
54        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetItemMoney)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
55        Err.Clear
End Function

Public Function funSampleCheckinInfoWrite(strAdvices As String, strName As String, strBatchNO As Long, Optional strErr As String, Optional ByVal strSentName As String) As Boolean
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����                   ǩ��ʱ�ѽ����˽���ʱ��д�뵽LIS��
      '
      '����                   strAdvices   ҽ����,���ҽ��ʹ��","�ŷָ�
      '                       strName      ǩ����
      '                       strBatchNO   ����
      '����                   �������Falseʱ����strErr����ʾ�������
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strValue As String
          Dim intloop As Integer
          Dim strSQL As String
          Dim strArr() As String
          Dim arrSql() As String
          Dim blnTran As Boolean

1         On Error GoTo funSampleCheckinInfoWrite_Error

2         If strAdvices <> "" Then
3             If Left(strAdvices, 1) = "|" Then strAdvices = Mid(strAdvices, 2)
4             strArr = Str2Array(strAdvices, "|", 4000)
5             ReDim arrSql(UBound(strArr))

6             For intloop = 0 To UBound(strArr)
7                 strValue = Replace(strArr(intloop), "|", ",")
8                 If VerCompare(gSysInfo.VersionLIS, "10.35.150") <> -1 Then
9                     strSQL = "Zl_��������ǩ��_Update('" & strValue & "','" & strName & "'," & IIf(strBatchNO = 0, "null", strBatchNO) & ",'" & strSentName & "',0,1)"
10                Else
11                    strSQL = "Zl_��������ǩ��_Update('" & strValue & "','" & strName & "'," & IIf(strBatchNO = 0, "null", strBatchNO) & ",'" & strSentName & "',0)"
12                End If
13                arrSql(intloop) = strSQL
14            Next

15            gcnLisOracle.BeginTrans
16            blnTran = True
17            For intloop = 0 To UBound(arrSql)
18                If arrSql(intloop) <> "" Then
19                    Call ComExecuteProc(Sel_Lis_DB, arrSql(intloop), "ǩ����Ϣ")
20                End If
21            Next
22            gcnLisOracle.CommitTrans
23            blnTran = False

24            funSampleCheckinInfoWrite = True
25        End If

          '����ˢ�¿��ڸſ�δ���ձ�ǩ����
26        Call SendMessage("RefreshDeptSurvey3")


27        Exit Function
funSampleCheckinInfoWrite_Error:
28        If blnTran Then gcnLisOracle.RollbackTrans
29        strErr = "ǩ��ʱ��д��־����" & Err.Number & " " & Err.Description
30        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funSampleCheckinInfoWrite)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
31        Err.Clear
End Function

Public Function Str2Array(ByVal strTxt As String, ByVal strDeli As String, ByVal intLength As Integer) As Variant
          '����: ������ָ�������ַ���,ת��������
          '����: strTxt     �����ַ���
          '      strDeli    �ָ��� - ��֧�ֵ��ַ��ָ���
          '      intLength  ָ����󳤶�
          Dim arrstr() As String
          
1         On Error GoTo Str2Array_Error
          
2         ReDim arrstr(0)
          
3         Do While Len(strTxt) >= intLength And Len(strDeli) = 1 And InStr(1, strTxt, strDeli) > 0
4             arrstr(UBound(arrstr)) = Left(strTxt, InStrRev(strTxt, strDeli, intLength) - 1)
5             strTxt = Mid(strTxt, Len(arrstr(UBound(arrstr))) + Len(strDeli) + 1)
6             If strTxt <> "" Then ReDim Preserve arrstr(UBound(arrstr) + 1)
7         Loop
          
8         If strTxt <> "" Then arrstr(UBound(arrstr)) = strTxt
          
9         Str2Array = arrstr
          
10        Exit Function
Str2Array_Error:
11        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(Str2Array)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
12        Err.Clear
End Function

Public Function funSampleSendInfo(strAdvices As String, intType As Integer, ByVal strUser As String, Optional strErr As String) As Boolean
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                   �걾�����˷���ʱ��д��LIS��
          '
          '����                   strAdvices   ҽ����,���ҽ��ʹ��","�ŷָ�
          '                       intType      --0Ϊ�ͼ� 1Ϊȡ���ͼ�
          '����                   �������Falseʱ����strErr����ʾ�������
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
          Dim strSQL As String
      '    SampleCheckinInfoWrite
1         On Error GoTo funSampleSendInfo_Error

2         strAdvices = Replace(strAdvices, "|", ",")
3         If Mid(strAdvices, Len(strAdvices), 1) = "," Then
4             strAdvices = Mid(strAdvices, 1, Len(strAdvices) - 1)
5         End If
6         strSQL = "Zl_���������ͼ�_Update('" & strAdvices & "'," & intType & ",'" & strUser & "')"
7         Call ComExecuteProc(Sel_Lis_DB, strSQL, "ǩ����Ϣ")
8         funSampleSendInfo = True
          
          '����ˢ�¿��ڸſ�δ�ǼǱ�ǩ����
9         Call SendMessage("RefreshDeptSurvey2")
          

10        Exit Function
funSampleSendInfo_Error:
11        strErr = "�걾�ͼ��д��־����" & Err.Number & " " & Err.Description
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funSampleSendInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
13        Err.Clear
End Function

Private Function InPatient() As ADODB.Recordset
    '��ʼ�����˼�¼��
    Dim rsRetur As New ADODB.Recordset
    
'    If rsRetur.State = adStateOpen Then rsRetur.Close
    rsRetur.Fields.Append "HIS����ID", adBigInt, 19
    rsRetur.Fields.Append "����", adVarChar, 100
    rsRetur.Fields.Append "�Ա�", adVarChar, 4
    rsRetur.Fields.Append "����", adVarChar, 20
    rsRetur.Fields.Append "��������", adVarChar, 4
    rsRetur.Fields.Append "���䵥λ", adVarChar, 10
    rsRetur.Fields.Append "������Դ", adVarChar, 4
    rsRetur.Fields.Append "����", adVarChar, 10
    rsRetur.Fields.Append "������", adVarChar, 20
    rsRetur.Fields.Append "���˿���", adVarChar, 100
    rsRetur.Fields.Append "�����", adVarChar, 19
    rsRetur.Fields.Append "סԺ��", adVarChar, 19
    rsRetur.Fields.Append "���˿��ұ���", adVarChar, 10
    rsRetur.Fields.Append "����", adVarChar, 100
    rsRetur.Fields.Append "��������", adVarChar, 100
    rsRetur.Fields.Append "��������", adVarChar, 18

    rsRetur.CursorLocation = adUseClient
    rsRetur.LockType = adLockOptimistic
    rsRetur.CursorType = adOpenStatic
    rsRetur.Open

    Set InPatient = rsRetur
End Function

Public Function funGetPatientAndAdivce(strFind As String, lngMachineType As Long, lngMachineID As Long, Optional strErr As String) As ADODB.Recordset
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����               ������Ĳ�����Ϣ����HIS������Ϣ�����ض�Ӧ�ļ�¼��
          '                   strFind = ���ҵ���Ϣ
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset, rsbuff As ADODB.Recordset
          Dim intloop As Integer
          Dim strAadvie As String, strbuff As String
          Dim varAdvices As Variant, strAdvice As String
          Dim blnGet As Boolean
          Dim strParentID As String
          Dim strBarCode As String

1         On Error GoTo funGetPatientAndAdivce_Error

2         If strFind = "" Then Exit Function
          
3         strSQL = "Select b.ҽ��id, b.��������, a.����id" & vbNewLine & _
                   "   From ����ҽ������ B, ����ҽ����¼ A" & vbNewLine & _
                   "   Where a.Id = b.ҽ��id And a.���id Is Not Null And  b.��������= [1]"
                  
4         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���˲���", strFind)
5         If rsTmp.RecordCount > 0 Then
6             Do Until rsTmp.EOF
7                 strAadvie = strAadvie & rsTmp("ҽ��id") & ","
8                 strBarCode = rsTmp("��������")
9                 strParentID = rsTmp("����id")
10                rsTmp.MoveNext
11            Loop
12            varAdvices = Split(strAadvie, ",")
13            For intloop = 0 To UBound(varAdvices)
14                If varAdvices(intloop) <> "" Then
15                    strbuff = "Select ����id,�������� from ����������� where ҽ��id =[1] "
16                    Set rsbuff = ComOpenSQL(Sel_Lis_DB, strbuff, "�ɼ�վ��ѯ", varAdvices(intloop))
17                    If rsbuff.EOF Then
                           '�����������뵥����
18                         strbuff = "Select Id,���id,�걾��λ,ִ�п���id From ����ҽ����¼  Where id=[1]"
19                         Set rsTmp = ComOpenSQL(Sel_His_DB, strbuff, "����ҽ����¼", varAdvices(intloop))
20                         If Val(rsTmp!���ID) <> 0 Then
21                             strAdvice = rsTmp!ID & "," & rsTmp!���ID & "," & rsTmp!ִ�п���id & "," & rsTmp!�걾��λ
22                             blnGet = SendLisApplication(strAdvice, "", strErr)
23                             If blnGet = False Then
'24                                 Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "����ҽ���������� ������ҽ��id��" & varAdvices(intloop) & "��SendLisApplication����δ���ɳɹ�", False)
25                             Else
'26                                 Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "����ҽ���������� ������ҽ��id��" & varAdvices(intloop) & "���ɳɹ�", False)
27                             End If
28                         End If
29                    Else
30                        If rsbuff("��������") & "" = "" Then
31                            strSQL = "Zl_������������_Update(2,'" & rsbuff("����id") & "','" & strBarCode & "')"
32                             Call ComExecuteProc(Sel_Lis_DB, strSQL, "д������")
33                        End If
34                    End If
                      
35                End If
36            Next
37            Set funGetPatientAndAdivce = GetPatientRecordCode(strBarCode, lngMachineType, lngMachineID, strErr)
38        Else
39            Set funGetPatientAndAdivce = rsTmp
40        End If
          
41        Exit Function
funGetPatientAndAdivce_Error:
42        strErr = "������(funGetPatientAndAdivce),������Ϣ:" & Err.Number & " " & Err.Description
43        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetPatientAndAdivce)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
44        Err.Clear
End Function

Private Function GetPatientRecordCode(strFind As String, lngMachineType As Long, lngMachineID As Long, Optional strErr As String) As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim strWhere As String
          Dim lngSampleID As Long
          Dim blnSelFind As Boolean
          Dim strBarCode As String

1         On Error GoTo GetPatientRecordCode_Error

2         If lngMachineType = 1 Then
              '΢����
3             strSQL = "select distinct a.His����id,����ID,nvl(a.Ӥ��,0) Ӥ�� ,decode(a.Ӥ��,null,a.����,0,a.����,a.Ӥ������ ) ����,decode(a.Ӥ��,null,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(a.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�," & vbNewLine & _
                          " decode(a.Ӥ��,null,a.����,0,a.����,null) ����,decode(a.Ӥ��,null,a.���䵥λ,0,a.���䵥λ,'Ӥ') ���䵥λ,decode(a.Ӥ��,null,a.��������,0,a.��������,null) ��������,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                          "from ����������� a,���������Ŀ b " & vbNewLine & _
                          "where a.���id = b.id  and nvl(a.����״̬,0) = 0  [����] " & vbNewLine & _
                          "union all " & vbNewLine & _
                          "select distinct a.His����id,����ID,nvl(a.Ӥ��,0) Ӥ�� ,decode(a.Ӥ��,null,a.����,0,a.����,a.Ӥ������ ) ����,decode(a.Ӥ��,null,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(a.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�," & vbNewLine & _
                          "   decode(a.Ӥ��,null,a.����,0,a.����,null) ����,decode(a.Ӥ��,null,a.���䵥λ,0,a.���䵥λ,'Ӥ') ���䵥λ,decode(a.Ӥ��,null,a.��������,0,a.��������,null) ��������,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                          "from ����������� a,���������Ŀ b " & vbNewLine & _
                          "where a.���id = b.id  and a.�걾id = [3]  [����] " & vbNewLine
          
4             If IsNumeric(strFind) And InStr("*-+./ABDG", Mid(strFind, 1, 1)) = 0 Then
                  '�Ȱ��������
5                 strWhere = " and �������� = [1] "
                  
6                 strSQL = Replace(strSQL, "[����]", strWhere)
7                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����", strFind, lngMachineID, lngSampleID)
8                 If rsTmp.RecordCount > 0 Then
9                     blnSelFind = False
10                    strBarCode = strFind
11                Else
12                    blnSelFind = True
13                End If
14            Else
15                blnSelFind = True
16            End If
              
17            If blnSelFind = True Then
18                blnSelFind = False
19                If (Left(strFind, 1) = "A" Or Left(strFind, 1) = "-") And IsNumeric(Mid(strFind, 2)) Then '����ID
20                    strWhere = " and a.HIS����ID = [4] "
21                    strFind = Mid(strFind, 2)
22                ElseIf (Left(strFind, 1) = "B" Or Left(strFind, 1) = "+") And IsNumeric(Mid(strFind, 2)) Then 'סԺ��
23                    strWhere = " and a.סԺ�� = [1] "
24                    strFind = Mid(strFind, 2)
25                ElseIf (Left(strFind, 1) = "D" Or Left(strFind, 1) = "*") And IsNumeric(Mid(strFind, 2)) Then '�����
26                    strWhere = " and a.����� = [1] "
27                    strFind = Mid(strFind, 2)
28                ElseIf Left(strFind, 1) = "G" Or Left(strFind, 1) = "." Then '�Һŵ�
29                    strWhere = " and a.�Һŵ� = [1] "
30                    strFind = Mid(strFind, 2)
31                ElseIf Left(strFind, 1) = "/" Then '�շѵ��ݺ�
32                    strWhere = " and a.�շѵ��� = [1] "
33                    strFind = Mid(strFind, 2)
34                Else
                      'û��ǰ׺ʱ�ڲ���id,סԺ��,�����,�Һŵ�,�շѵ��ݺ��в���
35                    strWhere = ""
                      
36                    strSQL = "Select Distinct B.His����id,����ID,nvl(b.Ӥ��,0) Ӥ�� , decode(b.Ӥ��,null,b.����,0,b.����,b.Ӥ������) ����,decode(b.Ӥ��,null,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(b.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�, " & vbNewLine & _
                              "   decode(b.Ӥ��,null,B.����,0,b.����,null) ����,decode(b.Ӥ��,null,b.���䵥λ,0,b.���䵥λ,'Ӥ') ���䵥λ,decode(b.Ӥ��,null,b.��������,0,b.��������,null) �������� ,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                              "From (Select His����id" & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where   His����id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where    סԺ�� =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where    ����� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where    �Һŵ� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His����id From ����������� Where   �շѵ��� = [1] ) A, ����������� B, ���������Ŀ C " & vbNewLine & _
                              "Where A.His����id = B.His����id And B.���id = C.Id  " & vbNewLine
                              
37                    strSQL = strSQL & " union all " & vbNewLine & _
                              "Select Distinct B.His����id,����ID ,nvl(b.Ӥ��,0) Ӥ�� ,decode(b.Ӥ��,null,b.����,0,b.����,b.Ӥ������) ����,decode(b.Ӥ��,null,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(b.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�, " & vbNewLine & _
                              "   decode(b.Ӥ��,null,B.����,0,b.����,null) ����,decode(b.Ӥ��,null,b.���䵥λ,0,b.���䵥λ,null) ���䵥λ,decode(b.Ӥ��,null,b.��������,0,b.��������,null) ��������,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                              "From (Select His����id" & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where �걾id = [3] and  His����id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  �걾id = [3] and  סԺ�� =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  �걾id = [3] and  ����� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  �걾id = [3] and  �Һŵ� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His����id From ����������� Where  �걾id = [3] and  �շѵ��� = [1] ) A, ����������� B, ���������Ŀ C " & vbNewLine & _
                              "Where A.His����id = B.His����id And B.���id = C.Id " & vbNewLine
38                End If
39                strSQL = Replace(strSQL, "[����]", strWhere)
40                strSQL = strSQL & strWhere
41                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����", strFind, lngMachineID, lngSampleID, Val(strFind))
42            End If
              
43        Else
              '��ͨ
44            strSQL = "select distinct a.His����id,����ID,nvl(a.Ӥ��,0) Ӥ�� ,decode(a.Ӥ��,null,a.����,0,a.����,a.Ӥ������ ) ����,decode(a.Ӥ��,null,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(a.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�," & vbNewLine & _
                          " decode(a.Ӥ��,null,a.����,0,a.����,null) ����,decode(a.Ӥ��,null,a.���䵥λ,0,a.���䵥λ,'Ӥ') ���䵥λ,decode(a.Ӥ��,null,a.��������,0,a.��������,null) ��������,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                          "from ����������� a,���������Ŀ b,�������ָ�� c,��������ָ�� d" & vbNewLine & _
                          "where a.���id = b.id and b.id = c.���id and c.��Ŀid = d.��Ŀid and nvl(a.����״̬,0) = 0 and d.����id = [2] [����] " & vbNewLine & _
                          "union all " & vbNewLine & _
                          "select distinct a.His����id,����ID,nvl(a.Ӥ��,0) Ӥ�� ,a.����,decode(a.�Ա�,1,'��',2,'Ů',9,'δ֪','������') �Ա�,a.����,a.���䵥λ,a.��������,decode(a.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                          "from ����������� a,���������Ŀ b,�������ָ�� c,��������ָ�� d" & vbNewLine & _
                          "where a.���id = b.id and b.id = c.���id and c.��Ŀid = d.��Ŀid and a.�걾id = [3] and d.����id = [2] [����] " & vbNewLine
          
45            If IsNumeric(strFind) And InStr("*-+./ABDG", Mid(strFind, 1, 1)) = 0 Then
                  '�Ȱ��������
46                strWhere = " and �������� = [1] "
47                lngSampleID = 0
48                strSQL = Replace(strSQL, "[����]", strWhere)
49                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����", strFind, lngMachineID, lngSampleID)
50                If rsTmp.RecordCount > 0 Then
51                    blnSelFind = False
52                    strBarCode = strFind
53                Else
54                    blnSelFind = True
55                End If
56            Else
57                blnSelFind = True
58            End If
59            If blnSelFind = True Then
60                If (Left(strFind, 1) = "A" Or Left(strFind, 1) = "-") And IsNumeric(Mid(strFind, 2)) Then '����ID
61                    strWhere = " and a.HIS����ID = [4] "
62                    strFind = Mid(strFind, 2)
63                ElseIf (Left(strFind, 1) = "B" Or Left(strFind, 1) = "+") And IsNumeric(Mid(strFind, 2)) Then 'סԺ��
64                    strWhere = " and a.סԺ�� = [1] "
65                    strFind = Mid(strFind, 2)
66                ElseIf (Left(strFind, 1) = "D" Or Left(strFind, 1) = "*") And IsNumeric(Mid(strFind, 2)) Then '�����
67                    strWhere = " and a.����� = [1] "
68                    strFind = Mid(strFind, 2)
69                ElseIf Left(strFind, 1) = "G" Or Left(strFind, 1) = "." Then '�Һŵ�
70                    strWhere = " and a.�Һŵ� = [1] "
71                    strFind = Mid(strFind, 2)
72                ElseIf Left(strFind, 1) = "/" Then '�շѵ��ݺ�
73                    strWhere = " and a.�շѵ��� = [1] "
74                    strFind = Mid(strFind, 2)
75                Else
                      'û��ǰ׺ʱ�ڲ���id,סԺ��,�����,�Һŵ�,�շѵ��ݺ��в���
76                    strWhere = ""
                      
77                    strSQL = "Select Distinct B.His����id,����ID,nvl(b.Ӥ��,0) Ӥ�� ,decode(b.Ӥ��,null,b.����,0,b.����,b.Ӥ������) ����, decode(b.Ӥ��,null,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(b.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�,  " & vbNewLine & _
                              "   decode(b.Ӥ��,null,B.����,0,b.����,null) ����,decode(b.Ӥ��,null,b.���䵥λ,0,b.���䵥λ,'Ӥ') ���䵥λ,decode(b.Ӥ��,null,b.��������,0,b.��������,null) ��������,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                              "From (Select His����id" & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  His����id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where    סԺ�� =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where   ����� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where    �Һŵ� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His����id From ����������� Where    �շѵ��� = [1] ) A, ����������� B, ���������Ŀ C, �������ָ�� D, ��������ָ�� E" & vbNewLine & _
                              "Where A.His����id = B.His����id And B.���id = C.Id And C.Id = D.���id And D.��Ŀid = E.��Ŀid  and e.����ID = [2] " & vbNewLine
                              
78                    strSQL = strSQL & " union all " & vbNewLine & _
                              "Select Distinct B.His����id,����ID ,nvl(b.Ӥ��,0) Ӥ�� ,decode(b.Ӥ��,null,b.����,0,b.����,b.Ӥ������ ) ����, decode(b.Ӥ��,null,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),0,decode(b.�Ա�,1,'��',2,'Ů',9,'δ֪','������'),decode(b.Ӥ���Ա�,1,'��',2,'Ů',9,'δ֪','������')) �Ա�," & vbNewLine & _
                              "   decode(b.Ӥ��,null,B.����,0,b.����,null) ����,decode(b.Ӥ��,null,b.���䵥λ,0,b.���䵥λ,'Ӥ') ���䵥λ,decode(b.Ӥ��,null,b.��������,0,b.��������,null) ��������,decode(b.������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ " & vbNewLine & _
                              "From (Select His����id" & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where �걾id = [3] and  His����id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  �걾id = [3] and  סԺ�� =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  �걾id = [3] and  ����� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His����id " & vbNewLine & _
                              "       From �����������" & vbNewLine & _
                              "       Where  �걾id = [3] and  �Һŵ� = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His����id From ����������� Where  �걾id = [3] and  �շѵ��� = [1] ) A, ����������� B, ���������Ŀ C, �������ָ�� D, ��������ָ�� E" & vbNewLine & _
                              "Where A.His����id = B.His����id And B.���id = C.Id And C.Id = D.���id And D.��Ŀid = E.��Ŀid  and e.����ID = [2] " & vbNewLine
79                End If
80                strSQL = Replace(strSQL, "[����]", strWhere)
81                strSQL = strSQL & strWhere
82                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����", strFind, lngMachineID, lngSampleID, Val(strFind))
83            End If
              
84        End If
              
85        If rsTmp.RecordCount = 0 Then
              '�ӵǼ��еĲ�������ȥ����
86            strSQL = "select Distinct HIS����ID,����ID,������, ����, �Ա�," & vbNewLine & _
                      "  ����,���䵥λ, ��������,decode(������Դ,1,'����',2,'סԺ',3,'Ժ��',4,'���','') ������Դ,����,�������� �ѱ�,decode(���˿���,null,�������,���˿���) ���˿���,�����,סԺ��, " & vbNewLine & _
                      " ��������,·��״̬,���˿��ұ���,����,�������� " & vbNewLine & _
                      " from ���鱨���¼ where ������=[1] "
87            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����", strFind)
              
88            If rsTmp.RecordCount = 0 Then blnSelFind = True
89        End If
90        Set GetPatientRecordCode = rsTmp

91        Exit Function
GetPatientRecordCode_Error:
92        strErr = "������(GetPatientRecordCode),������Ϣ:" & Err.Number & " " & Err.Description
93        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetPatientRecordCode)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
94        Err.Clear
End Function


Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer, Optional strErr As String) As String
      '���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
      '������intNum=��Ŀ���,Ϊ0ʱ�̶��������
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, intType As Integer
          Dim curDate As Date
          
1         On Error GoTo GetFullNO_Error

2         If Len(strNO) >= 8 Then
3             GetFullNO = Right(strNO, 8)
4             Exit Function
5         ElseIf Len(strNO) = 7 Then
6             GetFullNO = PreFixNO & strNO
7             Exit Function
8         ElseIf intNum = 0 Then
9             GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
10            Exit Function
11        End If
12        GetFullNO = strNO
          
13        strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=" & intNum
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ȡ����")
15        If Not rsTmp.EOF Then
16            intType = NVL(rsTmp!��Ź���, 0)
17            curDate = rsTmp!����
18        End If

19        If intType = 1 Then
              '���ձ��
20            strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
21            GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
22        Else
              '������
23            GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
24        End If


25        Exit Function
GetFullNO_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetFullNO)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear

End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function


Public Function GetPatiDiagnose(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int��Դ As Integer, Optional strErr As String) As String
      '���ܣ���ȡ����ָ���ξ�����������
      '������lng����ID=�Һ�ID����ҳID
      '      int��Դ=1-����,2-סԺ
      '���أ���"��"�ŷָ��Ķ����ϴ�
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
              
1         On Error GoTo GetPatiDiagnose_Error

2         strSQL = "Select ��¼��Դ,�������,��ϴ���,�������,�Ƿ�����,Mod(�������,10) as ���� From ������ϼ�¼" & _
              " Where ����ID=[1] And ��ҳID=[2] And ������� IN(" & IIf(int��Դ = 1, "1,11", "1,2,3,11,12,13") & ")" & _
              " Order by ��¼��Դ,�������,��ϴ���"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "GetPatiDiagnose", lng����ID, lng����ID)
          
          '�Ȱ���Դ����˳�����
4         rsTmp.Filter = "��¼��Դ=3" '��ҳ����
5         If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=2" '��Ժ�Ǽ�
6         If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=1" '����
7         If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=4" '������¼��
          
          'סԺ�ٰ���������˳�����
8         If Not rsTmp.EOF And int��Դ = 2 Then
9             strSQL = rsTmp.Filter
10            rsTmp.Filter = strSQL & " And ����=3"
11            If rsTmp.EOF Then rsTmp.Filter = strSQL & " And ����=2"
12            If rsTmp.EOF Then rsTmp.Filter = strSQL & " And ����=1"
13        End If
          
14        strSQL = ""
15        Do While Not rsTmp.EOF
16            If Not IsNull(rsTmp!�������) Then
17                strSQL = strSQL & "��" & rsTmp!������� & IIf(NVL(rsTmp!�Ƿ�����, 0) = 1, "������", "")
18            End If
19            rsTmp.MoveNext
20        Loop
          
21        GetPatiDiagnose = Mid(strSQL, 2)


22        Exit Function
GetPatiDiagnose_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetPatiDiagnose)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear

End Function

Public Function GetAppendItemValue(ByVal str��Ŀ As String, ByVal lngҪ��ID As Long, lng����ID As Long, _
             var����ID As Variant, strDiagnosis As String, intӤ�� As Integer, strAdvItem As String) As String
      '���ܣ���ȡָ�������븽��ֵ
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, strText As String
          Dim arrItem As Variant, i As Long
              
          '1.����ж�ӦҪ�أ���Ҫ����ȡ������ȡ
1         On Error GoTo GetAppendItemValue_Error

2         If lngҪ��ID <> 0 Then
3             If TypeName(var����ID) = "String" Then
4                 strSQL = "Select Zl_Replace_Element_Value(B.������,[1],A.ID,1) as ����" & _
                      " From ���˹Һż�¼ A,����������Ŀ B Where A.NO=[2] And B.ID=[3] And a.��¼����=1 And a.��¼״̬=1"
5                 Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS�ӿ�", lng����ID, CStr(var����ID), lngҪ��ID)
6             Else
7                 strSQL = "Select Zl_Replace_Element_Value(������,[1],[2],2) as ����" & _
                      " From ����������Ŀ Where ID=[3]"
8                 Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS�ӿ�", lng����ID, CStr(var����ID), lngҪ��ID)
9             End If
10            If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
11        End If
          
          '2.�����ϣ���δ�������¼���������ȡ
12        If str��Ŀ Like "*���" And strText = "" And strDiagnosis <> "" Then
13            strText = strDiagnosis
14        End If

          '3.δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰδ�����ҽ������ȡ,�������д��Ϊ׼
15        If strText = "" And strAdvItem <> "" Then
16            arrItem = Split(strAdvItem, "<Split1>")
17            For i = 0 To UBound(arrItem)
18                If Split(arrItem(i), "<Split2>")(0) = str��Ŀ Then
19                    strText = Split(arrItem(i), "<Split2>")(3): Exit For
20                End If
21            Next
22        End If
          
          '4.δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
23        If strText = "" Then
24            strSQL = _
                  " Select ���� From (" & _
                  " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
                  " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
                  IIf(TypeName(var����ID) = "String", " And A.�Һŵ�=[2]", " And A.��ҳID=[3]") & _
                  " And B.��Ŀ=[5] And B.���� is Not Null" & _
                  " Order by A.����ʱ�� Desc) Where Rownum=1"
25            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS�ӿ�", lng����ID, CStr(var����ID), Val(var����ID), intӤ��, str��Ŀ)
26            If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
27        End If
          
28        GetAppendItemValue = strText


29        Exit Function
GetAppendItemValue_Error:
30        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetAppendItemValue)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
31        Err.Clear

End Function

Public Function funWriteAdvicesLookState(strAdvices As String, intType As Integer, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                   д��ҽ���Ĳ���״̬
          '����                   strAdvices   ҽ����,���ҽ��ʹ��","�ŷָ�
          '                       intType      1=�Ѳ��� 0=δ����
          '����                   True=�ɹ�   False=ʧ��
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          
1         On Error GoTo funWriteAdvicesLookState_Error

2         strSQL = "Zl_�����������_update('" & strAdvices & "','" & intType & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "д�뱨���ѯ״̬")
          
4         funWriteAdvicesLookState = True


5         Exit Function
funWriteAdvicesLookState_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funWriteAdvicesLookState)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear

End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Currency
      '���ܣ���ȡָ�����˵��췢���ķ����ܶ�
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String

1         On Error GoTo GetPatiDayMoney_Error

2         strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlCISKernel", lng����ID)
4         If Not rsTmp.EOF Then GetPatiDayMoney = NVL(rsTmp!���, 0)


5         Exit Function
GetPatiDayMoney_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetPatiDayMoney)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear

End Function

Public Function ReCalcBirth(ByVal strOld As String, ByVal str���䵥λ As String) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����:              ������������䵥λ���㲡�˵ĳ�������,���䵥λΪ��ʱ,�������ռٶ�Ϊ1��1��,���䵥λΪ��ʱ,�������ڼٶ�Ϊ1��
          '
          '���:
          '                   strOld               ����
          '                   str���䵥λ          ���䵥λ
          '
          '����:              ��������
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strTmp As String, strFormat As String, lngDays As Long
          
1         On Error GoTo ReCalcBirth_Error

2         strTmp = "____-__-__"
3         If str���䵥λ = "" Then
4             strFormat = "YYYY-MM-DD"
5             If strOld Like "*��*��" Or strOld Like "*��*����" Then
6                 strFormat = "YYYY-MM-01"
7                 lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "��") + 1))
8             ElseIf strOld Like "*��*��" Or strOld Like "*����*��" Then
9                 lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "��") + 1))
10            ElseIf strOld Like "*��" Or IsNumeric(strOld) Then
11                strFormat = "YYYY-01-01"
12                lngDays = 365 * Val(strOld)
13            ElseIf strOld Like "*��" Or strOld Like "*����" Then
14                strFormat = "YYYY-MM-01"
15                lngDays = 30 * Val(strOld)
16            ElseIf strOld Like "*��" Then
17                lngDays = Val(strOld)
18            End If
19            If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, Currentdate), strFormat)
20        ElseIf strOld <> "" Then
21            Select Case str���䵥λ
                  Case "��"
22                    If Val(strOld) > 200 Then lngDays = -1
23                Case "��"
24                    If Val(strOld) > 2400 Then lngDays = -1
25                Case "��"
26                    If Val(strOld) > 73000 Then lngDays = -1
27            End Select
              
28            If lngDays = 0 Then
29                strTmp = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
30                strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, Currentdate), "YYYY-MM-DD")
                  
31                If str���䵥λ = "��" Then
32                    strTmp = Format(strTmp, "YYYY-01-01")
33                ElseIf str���䵥λ = "��" Then
34                    strTmp = Format(strTmp, "YYYY-MM-01")
35                End If
36            End If
37        End If
38        ReCalcBirth = strTmp


39        Exit Function
ReCalcBirth_Error:
40        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(ReCalcBirth)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
41        Err.Clear
End Function

Public Function funModifyApplyItemStateYJ(strAdvices As String, intType As Integer, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                   д��������Ŀ������״̬
          '����                   strAdvices   ҽ����,���ҽ��ʹ��","�ŷָ�
          '                       д��
          '����                   True=�ɹ�   False=ʧ��
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim strSQL As String
          
1         On Error GoTo funModifyApplyItemStateYJ_Error

2         strSQL = "Zl_�����������_Modify('" & strAdvices & "','" & intType & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "д�뱨���ѯ״̬")
          
4         funModifyApplyItemStateYJ = True


5         Exit Function
funModifyApplyItemStateYJ_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funModifyApplyItemStateYJ)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear

End Function

Public Function funModifyPathState(lngPartentID As Long, lngMainID As Long, lngPathSatae As Long, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                   д���ٴ�·����·��״̬
          '����                   lngPartentID   ����id
          '                       lngMainID      ��ҳid
          '                       lngPathSatae   ·��״̬
          '                       д��
          '����                   True=�ɹ�   False=ʧ��
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim strSQL As String
          
1         On Error GoTo funModifyPathState_Error

2         strSQL = "Zl_����·��״̬_Modify('" & lngPartentID & "','" & lngMainID & "','" & lngPathSatae & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "д�뱨���ѯ״̬")
          
4         funModifyPathState = True


5         Exit Function
funModifyPathState_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funModifyPathState)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear

End Function


'ZL9ComLib��ParseXMLToRecord����غ�������������ı䣬��ͬ���ı�
Public Function ParseXMLToRecord(ByVal strMsgNo As String, ByVal strXML As String) As ADODB.Recordset
      '���ܣ�����XML�ṹ���ַ�����ת���ɼ�¼������ʽ
          Dim rsMsg As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim objXML As Object
          Dim strTmp1 As String
          Dim strTmp2 As String
          
1         On Error GoTo ParseXMLToRecord_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         If InStr(",ZLHIS_EMR_021,ZLHIS_TRANSFUSION_001,ZLHIS_CHARGE_001,ZLHIS_PACS_005,ZLHIS_LIS_002,ZLHIS_LIS_003,ZLHIS_OPER_001," & _
              "ZLHIS_CIS_001,ZLHIS_CIS_002,ZLHIS_CIS_003,ZLHIS_CIS_004,ZLHIS_CIS_005,ZLHIS_CIS_015,", "," & strMsgNo & ",") = 0 Then Exit Function
          
          '�������Ϣ�����ܲ���������XML
4         Call objXML.OpenXMLDocument(IIf(InStr(strXML, "<message>") = 0, "<message>" & strXML & "</message>", strXML))
          
5         Set rsMsg = New ADODB.Recordset
6         rsMsg.Fields.Append "����ID", adBigInt
7         rsMsg.Fields.Append "����ID", adVarChar, 20
8         rsMsg.Fields.Append "�������ID", adBigInt
9         rsMsg.Fields.Append "���ﲡ��ID", adBigInt
10        rsMsg.Fields.Append "������Դ", adBigInt
11        rsMsg.Fields.Append "��Ϣ����", adVarChar, 4000
12        rsMsg.Fields.Append "���ѳ���", adVarChar, 8
13        rsMsg.Fields.Append "���ͱ���", adVarChar, 60
14        rsMsg.Fields.Append "ҵ���ʶ", adVarChar, 120
15        rsMsg.Fields.Append "���ȳ̶�", adBigInt
16        rsMsg.Fields.Append "�Ƿ�����", adBigInt
17        rsMsg.Fields.Append "�Ǽ�ʱ��", adVarChar, 60
18        rsMsg.Fields.Append "����IDs", adVarChar, 4000
19        rsMsg.Fields.Append "������Ա", adVarChar, 4000
20        rsMsg.CursorLocation = adUseClient
21        rsMsg.LockType = adLockOptimistic
22        rsMsg.CursorType = adOpenStatic
23        rsMsg.Open
          
24        rsMsg.AddNew
25        rsMsg!���ͱ��� = strMsgNo
26        rsMsg!�Ƿ����� = 0
27        rsMsg!���ȳ̶� = 1
28        rsMsg!������Ա = ""
29        rsMsg!����IDs = ""
          
30        Call objXML.GetSingleNodeValue("patient_id", strTmp1) '����id
31        Call objXML.GetSingleNodeValue("clinic_id", strTmp2) '����id
32        rsMsg!����ID = Val(strTmp1): strTmp1 = ""
33        rsMsg!����id = strTmp2
          
34        Call objXML.GetSingleNodeValue("send_time", strTmp1)
35        If strTmp1 <> "" Then
36            rsMsg!�Ǽ�ʱ�� = strTmp1
37        Else
38            rsMsg!�Ǽ�ʱ�� = Format(Currentdate, "yyyy-MM-dd HH:mm:ss")
39        End If

40        strTmp1 = "": strTmp2 = ""
41        If InStr(",ZLHIS_PACS_005,ZLHIS_LIS_002,ZLHIS_LIS_003,ZLHIS_OPER_001,", "," & strMsgNo & ",") > 0 Then
              
42            Call objXML.GetSingleNodeValue("clinic_dept_id", strTmp1)
43            Call objXML.GetSingleNodeValue("clinic_area_id", strTmp2)
              
             'LISϵͳ����ʱ��û�д�����id�����Ǳ���������ȡһ��
44            If (strMsgNo = "ZLHIS_LIS_002" Or strMsgNo = "ZLHIS_LIS_003") And Val(strTmp1) = 0 And Val(strTmp2) = 0 Then
45                strTmp1 = "": strTmp2 = ""
46                Call objXML.GetSingleNodeValue("clinic_dept_code", strTmp1)
47                Call objXML.GetSingleNodeValue("clinic_area_code", strTmp2)
48                If strTmp1 <> "" Or strTmp2 <> "" Then
49                    If strTmp1 = strTmp2 Then
50                        strSQL = "select id from ���ű� where ����=[1]"
51                        Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "ParseXMLToRecord", strTmp1)
52                        If Not rsTmp.EOF Then
53                            strTmp1 = rsTmp!ID
54                            strTmp2 = strTmp1
55                        Else
56                            strTmp1 = "": strTmp2 = ""
57                        End If
58                    Else
59                        strSQL = "select id,���� from ���ű� where ���� in (Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)))"
60                        Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "ParseXMLToRecord", strTmp1 & "," & strTmp2)
61                        If Not rsTmp.EOF Then
62                            rsTmp.Filter = "����='" & strTmp1 & "'"
63                            strTmp1 = IIf(Not rsTmp.EOF, rsTmp!ID, "")
64                            rsTmp.Filter = "����='" & strTmp2 & "'"
65                            strTmp2 = IIf(Not rsTmp.EOF, rsTmp!ID, "")
66                        Else
67                            strTmp1 = "": strTmp2 = ""
68                        End If
69                    End If
70                End If
71            End If
              
72            rsMsg!�������id = Val(strTmp1)
73            rsMsg!���ﲡ��id = Val(strTmp2)
74            If Val(strTmp2) <> Val(strTmp1) Then
75                rsMsg!����IDs = Val(strTmp1) & "," & Val(strTmp2)
76            Else
77                rsMsg!����IDs = Val(strTmp1)
78            End If
79            strTmp1 = ""
80            Call objXML.GetSingleNodeValue("patient_source", strTmp1)
81            rsMsg!������Դ = Val(strTmp1)
82            rsMsg!���ѳ��� = "0110"
              
83            strTmp1 = "": strTmp2 = ""
84            If strMsgNo = "ZLHIS_LIS_002" Then
85                rsMsg!��Ϣ���� = "�����ı��汻������"
86                rsMsg!���ѳ��� = "0100"
87                Call objXML.GetSingleNodeValue("specimen_id", strTmp1) '�걾id
88                rsMsg!ҵ���ʶ = Val(strTmp1)
89                rsMsg!���ȳ̶� = 2
90            ElseIf strMsgNo = "ZLHIS_LIS_003" Then
91                Call objXML.GetSingleNodeValue("element_title", strTmp1) 'Σ��ֵ����
92                Call objXML.GetSingleNodeValue("element_value", strTmp2) 'Σ��ֵֵ
93                rsMsg!��Ϣ���� = "Σ��ֵ��" & strTmp1 & "(" & strTmp2 & ")��"
                  
94                strTmp1 = ""
95                Call objXML.GetSingleNodeValue("order_id", strTmp1) 'ҽ��id
96                rsMsg!ҵ���ʶ = Val(strTmp1)
97                rsMsg!���ȳ̶� = 3
98            ElseIf strMsgNo = "ZLHIS_PACS_005" Then
99                Call objXML.GetSingleNodeValue("check_item_title", strTmp1) 'Σ��ֵֵ
100               rsMsg!��Ϣ���� = "Σ��ֵ��" & strTmp1 & "��"
101               Call objXML.GetSingleNodeValue("order_id", strTmp2) 'ҽ��id
102               rsMsg!ҵ���ʶ = Val(strTmp2)
103               rsMsg!���ȳ̶� = 3
104           ElseIf strMsgNo = "ZLHIS_OPER_001" Then
105               Call objXML.GetSingleNodeValue("operation_item_title", strTmp1) '��������
106               Call objXML.GetSingleNodeValue("operation_time", strTmp2) '����ʱ��
                  
107               strSQL = "select ���� from ���ű� where id=[1]"
108               Set rsTmp = OpenSQLRecord(strSQL, "ParseXMLToRecord", Val(rsMsg!�������id))
109               rsMsg!��Ϣ���� = rsTmp!���� & "��" & strTmp1 & "���ŵ���" & Format(strTmp2, "yyyy-MM-dd HH:mm")
                  
110               strTmp1 = "": strTmp2 = ""
111               Call objXML.GetSingleNodeValue("request_id", strTmp1) '����ҽ��id
112               rsMsg!ҵ���ʶ = Val(strTmp1)
113               Call objXML.GetSingleNodeValue("major_doctor", strTmp2) '����ҽʦ
114               rsMsg!������Ա = strTmp2
115           End If
116       End If
117       rsMsg.Update
          
118       If rsMsg.RecordCount > 0 Then
119           rsMsg.MoveFirst
120           Set ParseXMLToRecord = rsMsg
121       End If


122       Exit Function
ParseXMLToRecord_Error:
123       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(ParseXMLToRecord)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
124       Err.Clear
          
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
    '��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "�������Ӳ�������"
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function funExeDeptID(ByVal lngSampleID As Long) As Long
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����                   ��ȡִ�п���id
      '����
      '                       longSampleid        �걾id
      '
      '����                   ִ�п���id
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo funExeDeptID_Error

2         strSQL = "Select  d.Id ִ�п���id " & vbNewLine & _
                 "   From ���鱨���¼ A, ����������¼ B, ����С���¼ C, ���ű� D" & vbNewLine & _
                 "   Where a.����id = b.Id And b.С��id = c.Id And c.His���ű��� = d.���� and a.id =[1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ǩ��", lngSampleID)

4         If rsTmp.RecordCount > 0 Then
5             strSQL = "select Zl_Fun_Getsignpar(6," & rsTmp("ִ�п���ID") & ") as tag from dual "
6             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ǩ��")
7             If rsTmp.RecordCount > 0 Then
8                 funExeDeptID = rsTmp("tag")
9             End If
10        Else
11            funExeDeptID = 0
12        End If



13        Exit Function
funExeDeptID_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funExeDeptID)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear

End Function

Public Function ReviseDate(ByVal strDate As String) As String
'���ܣ���ʱ��ת��Ϊͳһ��24Сʱ��ʱ��
    If strDate = "" Then
        ReviseDate = ""
    Else
        ReviseDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
    End If
End Function

Public Function GetBabyInfor(ByVal lngPatientID As Long, ByVal lngPatientPage As Long, ByVal intBaby As Integer) As Recordset
      '���ܣ�����ĸ�׵Ĳ���ID ��ҳID �Լ�Ӥ�� ���غ��Ӽ�¼��
          Dim strSQL As String
          
1         On Error GoTo GetBabyInfor_Error

2         strSQL = "Select t.Ӥ������, t.Ӥ���Ա�, t.�������, t.���䷽ʽ, t.̥��״��,t.����ʱ��," & vbNewLine & _
                  "Nvl(Round(Nvl(t.����ʱ��, Sysdate) - t.����ʱ��), 0) ||'��' As ����,t.��� Ӥ�����" & vbNewLine & _
                  "From ������������¼ t" & vbNewLine & _
                  "Where t.����id = [1] And t.��ҳid = [2] And t.��� =[3]"
3         Set GetBabyInfor = ComOpenSQL(Sel_His_DB, strSQL, "��������", lngPatientID, lngPatientPage, intBaby)


4         Exit Function
GetBabyInfor_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetBabyInfor)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
6         Err.Clear
          
End Function

Public Function funGetLabNewReportList(ByVal lngPatientID As Long, ByVal lngMainID As Long, ByRef strXMLNewLIS As String, Optional lngApplyID As Long) As Boolean
          '����               LIS�Ĺ�����������������XML��ʽ�Ĳ��˵ļ��鱨���б�
          '����
          '                   lngPatientID            ��Ϣͷ
          '                   lngPatientID            ��Ϣ����
          '                   strXMLNewLIS            ���ص��ִ�
          '                   lngApplyID              ����id
          '����               True=�ɹ�   False=ʧ��
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim objXML As Object
          Dim i As Long

1         On Error GoTo funGetLabNewReportList_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         strSQL = "Select b.id ���鱨��id,a.����id ,a.���� ������־,c.���� ������Ŀ,b.�걾���,b.΢���� �Ƿ�΢����,0 ������,b.������,b.�����,b.���ʱ��,a.����ʱ�� " & vbNewLine & _
                   "   from ����������� a,���鱨���¼ b,���������Ŀ c where a.�걾id= b.id and a.���id = c.id  and  a.����id = [1] and a.��ҳid =[2] and a.����id is not null"
                   
4         If lngApplyID > 0 Then
5             strSQL = strSQL & " and ����id =[3]"
6         End If
7         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���˼��鱨���б�", lngPatientID, lngMainID, lngApplyID)
8         If rsTmp.RecordCount > 0 Then
9             With objXML
10                .ClearXmlText
11                .AppendNode "���鱨���б�" ', True '���ڵ�[���鱨���б�]
12                For i = 1 To rsTmp.RecordCount
13                    .AppendData "���鱨��id", rsTmp!���鱨��id '<���鱨��id>���ͣ�
14                    .AppendData "����id", rsTmp!����id '<����id>���ͣ�
15                    .AppendData "������־", rsTmp!������־ & "" '<������־>���ͣ�
16                    .AppendData "������Ŀ", rsTmp!������Ŀ & ""  '<������Ŀ>���ͣ�
17                    .AppendData "�걾���", rsTmp!�걾��� '<�걾���>���ͣ�
18                    .AppendData "�Ƿ�΢����", rsTmp!�Ƿ�΢���� & ""  '<΢����걾>���ͣ�
19                    .AppendData "������", Val(rsTmp!������ & "") '<������>���ͣ�
20                    .AppendData "������", rsTmp!������ & "" '<������>���ͣ�
21                    .AppendData "�����", rsTmp!����� & ""  '<�����>���ͣ�
22                    .AppendData "���ʱ��", rsTmp!���ʱ�� & "" '<���ʱ��>���ͣ�
23                    .AppendData "����ʱ��", rsTmp!����ʱ�� & ""  '<����ʱ��>���ͣ�
24                    rsTmp.MoveNext
25                Next
26                .AppendNode "���鱨���б�", True
27                If strXMLNewLIS = "" Then strXMLNewLIS = .XmlText
28            End With
29        End If
30        funGetLabNewReportList = True


31        Exit Function
funGetLabNewReportList_Error:
32        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetLabNewReportList)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
33        Err.Clear

End Function

Public Function funGetLabNewReportResultList(ByVal lngRepottID As Long, ByRef strXMLOldLIS As String) As Boolean
          '���ܣ�                 LIS�Ĺ�����������ȡ���˵ļ��鱨����
          '����
          'lngRepottID            ����id
          'strXMLOldLIS           ���ص��ִ�
          '����                   XML��ʽ���ִ�
          Dim strSQL As String
          Dim rsNewTmp As ADODB.Recordset
          Dim objXML As Object
          Dim strBH As String
          Dim i As Long

          '���°�
1         On Error GoTo funGetLabNewReportResultList_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         strSQL = "select  id,΢���� from ���鱨���¼ where id = [1]"
4         Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���˼������б�", lngRepottID)
5         If rsNewTmp.RecordCount > 0 Then
6             If Val(rsNewTmp("΢����") & "") = 1 Then
7                 strSQL = "Select distinct a.ϸ��id, b.������ ϸ����, a.�������� ����, a.��ҩ����, e.������ ������, c.��� �����ؽ��, c.������� ��ҩ��, c.ҩ������, e.�÷�����1, e.�÷�����2, e.ѪҩŨ��1," & vbNewLine & _
                           "          e.ѪҩŨ��2 , e.��ҩŨ��1, e.��ҩŨ��2" & vbNewLine & _
                           "   From ���鱨��ϸ�� A, ����ϸ����¼ B, ���鱨��ҩ�� C, ����ҩ�� E" & vbNewLine & _
                           "   Where a.ϸ��id = b.Id And b.Id = c.ϸ��id And c.ҩ��id = e.Id and a.�걾id=[1] order by b.������"
                  
8                 Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���˼������б�", lngRepottID)
9                 If rsNewTmp.RecordCount > 0 Then
10                    With objXML
11                        .ClearXmlText
12                        .AppendNode "΢������Ŀ" ', True '���ڵ�[��ͨ��Ŀ]
13                        For i = 1 To rsNewTmp.RecordCount
14                            If strBH <> rsNewTmp!ϸ���� & "" Then
15                                If strBH <> "" Then
16                                    .AppendNode "�����ؽ���б�", True
17                                End If
18                                strBH = rsNewTmp!ϸ���� & ""
19                                .AppendData "ϸ��id", rsNewTmp!ϸ��id & "" '<ϸ��id>���ͣ�
20                                .AppendData "ϸ����", rsNewTmp!ϸ���� & "" '<ϸ����>���ͣ�
21                                .AppendData "����", rsNewTmp!���� & "" '<����>���ͣ�
22                                .AppendData "��ҩ����", rsNewTmp!��ҩ���� & ""  '<��ҩ����>���ͣ�
23                                .AppendNode "�����ؽ���б�"  ', True '���ڵ�[ָ������]
24                            End If
                          
25                            .AppendData "������", rsNewTmp!������ & "" '<������>���ͣ�
26                            .AppendData "�����ؽ��", rsNewTmp!�����ؽ�� & "" '<�����ؽ��>���ͣ�
27                            .AppendData "��ҩ��", rsNewTmp!��ҩ�� & "" '<��ҩ��>���ͣ�
28                            .AppendData "ҩ������", rsNewTmp!ҩ������ & ""  '<ҩ������>���ͣ�
29                            .AppendData "�÷�����1", rsNewTmp!�÷�����1 & "" '<�÷�����1>���ͣ�
30                            .AppendData "�÷�����2", rsNewTmp!�÷�����2 & ""  '<�÷�����2>���ͣ�
31                            .AppendData "ѪҩŨ��1", rsNewTmp!ѪҩŨ��1 & "" '< ѪҩŨ��1 > ����:
32                            .AppendData "ѪҩŨ��2", rsNewTmp!ѪҩŨ��2 & "" '<ѪҩŨ��2>���ͣ�
33                            .AppendData "��ҩŨ��1", rsNewTmp!��ҩŨ��1 & ""  '<��ҩŨ��1>���ͣ�
34                            .AppendData "��ҩŨ��2", rsNewTmp!��ҩŨ��2 & ""  '<��ҩŨ��2>���ͣ�
35                            rsNewTmp.MoveNext
36                        Next
37                        .AppendNode "�����ؽ���б�", True
38                        .AppendNode "΢������Ŀ", True
39                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
40                    End With
41                End If
42            Else
43                strSQL = "Select a.��Ŀid ָ��id, b.ָ����� ָ�����, b.Ӣ���� ָ��Ӣ����, b.������ ָ��������," & vbNewLine & _
                           " a.������, a.�����־, a.����ο�, a.�������, b.��˽��Ŀ,a.��λ " & vbNewLine & _
                           "   From ���鱨����ϸ A, ����ָ�� B" & vbNewLine & _
                           "   Where a.��Ŀid = b.Id and a.�걾id =[1] "
44                Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���˼������б�", lngRepottID)
45                If rsNewTmp.RecordCount > 0 Then
46                    With objXML
47                        .ClearXmlText
48                        .AppendNode "��ͨ��Ŀ" ', True '���ڵ�[��ͨ��Ŀ]
49                        .AppendNode "ָ������" ', True '���ڵ�[ָ������]
50                        For i = 1 To rsNewTmp.RecordCount
51                            .AppendData "ָ��id", rsNewTmp!ָ��id & "" '<ָ��id>���ͣ�
52                            .AppendData "ָ�����", rsNewTmp!ָ����� & "" '<ָ�����>���ͣ�
53                            .AppendData "ָ��Ӣ����", rsNewTmp!ָ��Ӣ���� & "" '<ָ��Ӣ����>���ͣ�
54                            .AppendData "ָ��������", rsNewTmp!ָ�������� & ""  '<ָ��������>���ͣ�
55                            .AppendData "������", rsNewTmp!������ & ""  '<������>���ͣ�
56                            .AppendData "�����־", rsNewTmp!�����־ & ""  '<�����־>���ͣ�
57                            .AppendData "����ο�", rsNewTmp!����ο� & "" '< ����ο� > ����:
58                            .AppendData "�������", rsNewTmp!������� & "" '<�������>���ͣ�
59                            .AppendData "��˽��Ŀ", rsNewTmp!��˽��Ŀ & ""  '<��˽��Ŀ>���ͣ�
60                            .AppendData "��λ", rsNewTmp!��λ & ""          '<��λ>���ͣ��ַ�
61                            rsNewTmp.MoveNext
62                        Next
63                        .AppendNode "ָ������", True
64                        .AppendNode "��ͨ��Ŀ", True
65                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
66                    End With
67                End If
68            End If
69        End If
70        funGetLabNewReportResultList = True


71        Exit Function
funGetLabNewReportResultList_Error:
72        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetLabNewReportResultList)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
73        Err.Clear

End Function

Public Function funGetNewBloodBankItem(ByVal lngApplyID As Long, ByRef strXMLOldLIS As String) As Boolean
          '���ܣ�                 LIS�Ĺ�����������ȡ���˵ļ��鱨����
          '����
          'lngApplyID            ���id ,ҽ��id
          '����                   XML��ʽ���ִ�
          
          Dim strSQL As String
          Dim rsNewTmp As ADODB.Recordset
          Dim objXML As Object
          Dim i As Long
          
          '���°�
1         On Error GoTo funGetNewBloodBankItem_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         strSQL = "Select a.��Ŀid ָ��id, b.ָ����� ָ�����, b.Ӣ���� ָ��Ӣ����, b.������ ָ��������, a.������, a.�����־, a.����ο�" & vbNewLine & _
                   "   From ���鱨����ϸ A, ����ָ�� B,���鱨���¼ c,����������� d " & vbNewLine & _
                   "   Where a.��Ŀid = b.Id and a.�걾id = c.id  and  c.id = d.�걾id  and d.����id =[1] "
4         Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���˼������б�", lngApplyID)
5         If rsNewTmp.RecordCount > 0 Then
6             With objXML
7                 .ClearXmlText
8                 .AppendNode "��ͨ��Ŀ" ', True '���ڵ�[��ͨ��Ŀ]
9                 .AppendNode "ָ������" ', True '���ڵ�[ָ������]
10                For i = 1 To rsNewTmp.RecordCount
11                    .AppendData "ָ��id", rsNewTmp!ָ��id & "" '<ָ��id>���ͣ�
12                    .AppendData "ָ�����", rsNewTmp!ָ����� & "" '<ָ�����>���ͣ�
13                    .AppendData "ָ��Ӣ����", rsNewTmp!ָ��Ӣ���� & "" '<ָ��Ӣ����>���ͣ�
14                    .AppendData "ָ��������", rsNewTmp!ָ�������� & ""  '<ָ��������>���ͣ�
15                    .AppendData "������", rsNewTmp!������ & ""  '<������>���ͣ�
16                    .AppendData "�����־", rsNewTmp!�����־ & ""  '<�����־>���ͣ�
17                    .AppendData "����ο�", rsNewTmp!����ο� & "" '<����ο�> ����:
18                    rsNewTmp.MoveNext
19                Next
20                .AppendNode "ָ������", True
21                .AppendNode "��ͨ��Ŀ", True
22                If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
23            End With
24        End If
25        funGetNewBloodBankItem = True


26        Exit Function
funGetNewBloodBankItem_Error:
27        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetNewBloodBankItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
28        Err.Clear

End Function


Public Function funGetNewTransFusionApplyFor(strItemCodeing As String, lngPatientID As Long, intPatientType As Integer, lngHomePageID As Long, Optional strRegistrationBill As String, _
                                             Optional intBaby As Integer, Optional intType As Integer, Optional ByVal intDay As Integer) As String
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '����                   LIS�Ĺ������������ݴ���ҽ��ID���ؽ��
      '����
      '                       strItemCodeing ������Ŀ���루�ɴ�������ʹ�ö��ŷָ���
      '                       lngPatientID ����ID
      '                       intPatientType ������Դ 1-���2-סԺ
      '                       lngHomePageID ��ҳID ��������Դ=2ʱ��ѯ)
      '                       lngRegistrationBill �Һŵ�NO��������Դ<>2ʱ��ѯ���ξ��
      '                       intBaby           �Ƿ�Ӥ��
      '                       intType           ���ַ�ʽ��1=�ٴ˲�7���ڵġ�0 = ����ѯ 2=ָ����ѯ������intDay����������= �ݶ�
      '                       intDay            ��intType=2ʱ���˲�������Ч��ָ��Ҫ��ѯ�����������
      '�걾��ɸ�ʽ
      '                   ָ��1<split1>���Ʊ���1<split1>��λ1<split1>��˽��Ŀ1<split1>ָ�����1<split1>������1<split1>Ӣ����1<split1>ȡֵ����1<split1>
      '                       ������1<split2>�����־1<split2>�������1<split2>�������1<split2>�걾����1<split3>
      '                   ָ��2<split1>���Ʊ���2<split1>��λ2<split1>��˽��Ŀ2<split1>ָ�����2<split1>������2<split1>Ӣ����2<split1>ȡֵ����2<split1>
      '                       ������2<split2>�����־2<split2>�������2<split2>�������2<split2>�걾����2<split3>
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim rsTmpRuest As New ADODB.Recordset
          Dim strItemcodeOne As String
          Dim strSampleOne As String, i As Integer
          Dim strSampleTwo As String
          Dim varItemCodeing As Variant
          Dim strBH As String, strGetSeques As String
          Dim strStartTime As String
          Dim strEndTime As String

1         On Error GoTo funGetNewTransFusionApplyFor_Error

2         strEndTime = Format(Currentdate, "yyyy-mm-dd 23:59:59")
3         If intType = 1 Then
4             strStartTime = Format(Currentdate - 7, "yyyy-mm-dd 00:00:00")
5         ElseIf intType = 2 Then
6             strStartTime = Format(Currentdate - intDay, "yyyy-mm-dd 00:00:00")
7         End If
          '�ָ��ĳ���
          Const conSplit1 As String = "<split1>"                        '���ڷָ��걾,ʹ�á�<split1>���ָ�
          Const conSplit2 As String = "<split2>"                        '���ڷָ��걾��Ϣ,ʹ�á�<split2>���ָ�
          Const conSplit3 As String = "<split3>"                        '���ڷָ��걾ָ����Ϣ,ʹ�á�<split3>���ָ�
          Const conSplit4 As String = "<split4>"                        '���ڷָ�ָ������Ϣ,ʹ�á�<split4>���ָ�

          '--------------------------------------------------------------------------------------------------------------------------------------------------------------
8         varItemCodeing = Split(strItemCodeing, ",")
9         For i = LBound(varItemCodeing) To UBound(varItemCodeing)
10            strItemcodeOne = varItemCodeing(i)
11            If gUserInfo.NodeNo <> "-" Then
12                strSQL = "Select distinct d.id ָ��id, d.������ || '(' || d.Ӣ���� || ')' ָ��, d.��λ, 0 ��˽��Ŀ, d.ָ�����, d.������, d.Ӣ����, f.ȡֵ����" & vbNewLine & _
                         "   From �������ָ�� A, ����ָ�� D, ���������Ŀ E, ��������ָ�� F" & vbNewLine & _
                         "   Where a.��Ŀid = d.Id And d.Id = f.��Ŀid And a.���id = e.Id And e.���Ʊ��� = [1] and (e.վ��=[2] or e.վ�� is null)" & vbNewLine & _
                         "   order by d.id "
13            Else
14                strSQL = "Select distinct d.id ָ��id, d.������ || '(' || d.Ӣ���� || ')' ָ��, d.��λ, 0 ��˽��Ŀ, d.ָ�����, d.������, d.Ӣ����, f.ȡֵ����" & vbNewLine & _
                         "   From �������ָ�� A, ����ָ�� D, ���������Ŀ E, ��������ָ�� F" & vbNewLine & _
                         "   Where a.��Ŀid = d.Id And d.Id = f.��Ŀid And a.���id = e.Id And e.���Ʊ��� = [1]" & vbNewLine & _
                         "   order by d.id "
15            End If
16            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", strItemcodeOne, gUserInfo.NodeNo)
17            strBH = "***"
18            Do Until rsTmp.EOF

19                If strBH <> rsTmp("ָ��id") Then
20                    If strBH <> "***" Then
21                        If strSampleTwo = "��" Then
22                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & ""
23                        Else
24                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & strSampleTwo
25                        End If
26                        strSampleTwo = ""
27                    End If
                      '                If strBH = "***" Then
28                    strSampleOne = strSampleOne & conSplit3 & rsTmp("ָ��") & conSplit1 & strItemcodeOne & conSplit1 & rsTmp("��λ") & _
                                     conSplit1 & rsTmp("��˽��Ŀ") & conSplit1 & rsTmp("ָ�����") & _
                                     conSplit1 & rsTmp("������") & conSplit1 & rsTmp("Ӣ����") & conSplit1
                      '                Else
                      '                     strSampleOne = strSampleOne & strGetSeques & conSplit1 & conSplit3 & rsTmp("ָ��") & conSplit1 & strItemcodeOne & conSplit1 & rsTmp("��λ") & _
                                            '                                    conSplit1 & rsTmp("��˽��Ŀ") & conSplit1 & rsTmp("ָ�����") & _
                                            '                                    conSplit1 & rsTmp("������") & conSplit1 & rsTmp("Ӣ����") & conSplit1
                      '                End If
29                    strSQL = " Select *" & vbNewLine & _
                             "    From (Select b.���ʱ��, c.������, Decode(c.�����־, 1, '', 2, '��', 3, '��', 4, '�쳣', 5, '����', 6, '����', '') �����־, c.����ο�, c.�������," & vbNewLine & _
                             "                  B.�걾����" & vbNewLine & _
                             "           From ����������� A, ���鱨���¼ B, ���鱨����ϸ C, ����ָ�� D" & vbNewLine & _
                             "           Where a.�걾id = b.Id And b.Id = c.�걾id And c.��Ŀid = d.Id And Nvl(b.΢����, 0) <> 1 And a.���id = c.���id And" & vbNewLine & _
                             "                  b.���ʱ�� Is Not Null [����]  and d.Id =[6] and b.������Դ=[5] order by b.���ʱ�� desc ) E" & vbNewLine & _
                             "    Where Rownum = 1"
30                    If intPatientType = 2 Then
31                        If intBaby <> 0 Then
32                            strSQL = Replace(strSQL, "[����]", " and b.HIS����ID = [1] and b.��ҳid=[2]   and  a.Ӥ�� =[7] ")
33                        Else
34                            strSQL = Replace(strSQL, "[����]", " and b.HIS����ID = [1] and b.��ҳid=[2]   and nvl(a.Ӥ��,0)= 0 ")
35                        End If
36                        Set rsTmpRuest = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatientID, lngHomePageID, strRegistrationBill, strItemcodeOne, intPatientType, Val(rsTmp("ָ��id") & ""), intBaby)
37                        If rsTmpRuest.RecordCount > 0 Then
38                            strSampleTwo = rsTmpRuest("������") & conSplit2 & rsTmpRuest("�����־") & conSplit2 & rsTmpRuest("����ο�") & conSplit2 & rsTmpRuest("�������") & conSplit2 & rsTmpRuest("�걾����")
39                        Else
40                            If intType = 1 Or intType = 2 Then
41                                strSQL = " Select *" & vbNewLine & _
                                         "    From (Select b.���ʱ��, c.������, Decode(c.�����־, 1, '', 2, '��', 3, '��', 4, '�쳣', 5, '����', 6, '����', '') �����־, c.����ο�, c.�������," & vbNewLine & _
                                         "                  B.�걾����" & vbNewLine & _
                                         "           From ����������� A, ���鱨���¼ B, ���鱨����ϸ C, ����ָ�� D" & vbNewLine & _
                                         "           Where a.�걾id = b.Id And b.Id = c.�걾id And c.��Ŀid = d.Id And Nvl(b.΢����, 0) <> 1 And a.���id = c.���id And" & vbNewLine & _
                                         "                  b.���ʱ�� Is Not Null [����]  and d.Id =[5] order by b.���ʱ�� desc ) E" & vbNewLine & _
                                         "    Where Rownum = 1"
42                                If intBaby <> 0 Then
43                                    strSQL = Replace(strSQL, "[����]", " and b.HIS����ID = [1] and b.���ʱ�� between [2] and [3]  and a.Ӥ��=[6] ")
44                                Else
45                                    strSQL = Replace(strSQL, "[����]", " and b.HIS����ID = [1] and b.���ʱ�� between [2] and [3]   and nvl(a.Ӥ��,0)= 0 ")
46                                End If
47                                Set rsTmpRuest = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatientID, CDate(strStartTime), CDate(strEndTime), strItemcodeOne, Val(rsTmp("ָ��id") & ""), intBaby)
48                                If rsTmpRuest.RecordCount > 0 Then
49                                    strSampleTwo = rsTmpRuest("������") & conSplit2 & rsTmpRuest("�����־") & conSplit2 & rsTmpRuest("����ο�") & conSplit2 & rsTmpRuest("�������") & conSplit2 & rsTmpRuest("�걾����")
50                                Else
51                                    strSampleTwo = "��"
52                                End If
53                            Else
54                                strSampleTwo = "��"
55                            End If
56                        End If
57                    Else
58                        strSQL = Replace(strSQL, "[����]", " and b.HIS����ID = [1] and  b.�Һŵ�=[3] ")
59                        Set rsTmpRuest = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ���", lngPatientID, lngHomePageID, strRegistrationBill, strItemcodeOne, intPatientType, Val(rsTmp("ָ��id") & ""))
60                        If rsTmpRuest.RecordCount > 0 Then
61                            strSampleTwo = rsTmpRuest("������") & conSplit2 & rsTmpRuest("�����־") & conSplit2 & rsTmpRuest("����ο�") & conSplit2 & rsTmpRuest("�������") & conSplit2 & rsTmpRuest("�걾����")
62                        Else
63                            strSampleTwo = "��"
64                        End If
65                    End If
66                    strBH = rsTmp("ָ��id")
67                    strGetSeques = rsTmp("ȡֵ����") & ""
68                Else
69                    If strGetSeques <> "" Then
70                        strGetSeques = GetSameString(strGetSeques & "," & rsTmp("ȡֵ����"))
71                        If strSampleTwo = "��" Then
72                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & ""
73                        Else
74                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & strSampleTwo
75                        End If
76                        strSampleTwo = ""
77                        strGetSeques = ""
78                    End If
79                End If
80                rsTmp.MoveNext
81            Loop
82            If strSampleTwo = "��" Then
83                strSampleOne = strSampleOne & strGetSeques & conSplit1 & ""
84                strSampleTwo = ""
85                strGetSeques = ""
86            Else
87                strSampleOne = strSampleOne & strGetSeques & conSplit1 & strSampleTwo
88                strSampleTwo = ""
89                strGetSeques = ""
90            End If

91        Next
          '------------------------------------------------------------------------------------------------------------------------
92        If strSampleOne <> "" Then
93            strSampleOne = Mid(strSampleOne, Len(conSplit1) + 1)
94        End If
95        funGetNewTransFusionApplyFor = strSampleOne


96        Exit Function
funGetNewTransFusionApplyFor_Error:
97        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetNewTransFusionApplyFor)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
98        Err.Clear

End Function

Public Function GetSameString(ByVal strGetSeque As String) As String
          '����ظ�ȡֵ����
          Dim i As Integer
          Dim varGetSeque As Variant
          Dim strGetSeques As String

1         On Error GoTo GetSameString_Error

2         varGetSeque = Split(strGetSeque, ",")
3         For i = LBound(varGetSeque) To UBound(varGetSeque)
4             If varGetSeque(i) <> "" Then
5                 strGetSeques = "," & varGetSeque(i) & ","
6             Else
7                 strGetSeques = ""
8             End If
9             If Left(strGetSeque, 1) <> "," Then strGetSeque = "," & strGetSeque
10            If Right(strGetSeque, 1) <> "," Then strGetSeque = strGetSeque & ","
11            If InStr(GetSameString, strGetSeques) = 0 Then
12                GetSameString = GetSameString & varGetSeque(i) & ","
13            End If
14        Next
15        If GetSameString <> "" Then GetSameString = Mid(GetSameString, 1, Len(GetSameString) - 1)


16        Exit Function
GetSameString_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetSameString)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear

End Function

Public Function funFindAdvicePay(ByVal strAdvice As String, ByVal intPaentType As Integer, Optional ByVal strErr As String = "") As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim varAdvice As Variant
          Dim intloop As Integer
          Dim strAdvivePay As String
          Dim blnNewPait As Boolean       '�����Ƿ�Ϊ�����ﲡ��
          Dim intNewPaitPay As Integer    '�����ﲡ���Ƿ��շ�

1         On Error GoTo funFindAdvicePay_Error

2         varAdvice = Split(strAdvice, ",")
3         For intloop = 0 To UBound(varAdvice)

              '��������ﲡ���Ƿ��շ�
4             If gcnHisOracle.State = 1 Then
5                 If VerCompare(gSysInfo.VersionHIS, "10.35.100") <> -1 Then
6                     blnNewPait = funNewSystemSvr(Val(varAdvice(intloop)))
7                 End If
8             End If
9             If blnNewPait Then  '�����ﲡ��
                  '������
10                Exit Function
11            Else

                  '�������ﲡ��
12                If GetAdviceFeeKind(Val(varAdvice(intloop))) = 2 Then
13                    strSQL = "Select ҽ�����,��¼״̬,��¼����,ʵ�ս�� from סԺ���ü�¼ t where  t.ҽ����� in  (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)))"
14                Else
15                    strSQL = "Select ҽ�����,��¼״̬,��¼����,ʵ�ս�� from ������ü�¼ t where  t.ҽ����� in  (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) "
16                End If
17                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��ѯ�걾ID", varAdvice(intloop))
18                rsTmp.Filter = " ��¼����= 2 and ��¼״̬=1"
19                If rsTmp.RecordCount > 0 Then
20                    If Val(rsTmp("ʵ�ս��")) = 0 Then
21                        strAdvivePay = strAdvivePay & varAdvice(intloop) & ",-1|"
22                    Else
23                        strAdvivePay = strAdvivePay & varAdvice(intloop) & ",3|"
24                    End If
25                Else
26                    rsTmp.Filter = "��¼״̬=0"
27                    If rsTmp.RecordCount > 0 Then
28                        strAdvivePay = strAdvivePay & varAdvice(intloop) & ",0|"
29                    Else
30                        rsTmp.Filter = "��¼״̬=1"
31                        If rsTmp.RecordCount > 0 Then
32                            If Val(rsTmp("ʵ�ս��")) = 0 Then
33                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",-1|"
34                            Else
35                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",1|"
36                            End If
37                        Else
38                            rsTmp.Filter = "��¼״̬=2 or ��¼״̬=3"
39                            If rsTmp.RecordCount > 0 Then
40                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",2|"
41                            Else
42                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",0|"
43                            End If
44                        End If
45                    End If
46                End If
47            End If
48        Next
49        If strAdvivePay <> "" Then funFindAdvicePay = Mid(strAdvivePay, 1, Len(strAdvivePay) - 1)

50        Exit Function
funFindAdvicePay_Error:
51        strErr = "����funFindAdvicePay����" & Err.Number & " " & Err.Description
52        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funFindAdvicePay)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
53        Err.Clear
End Function

Private Function GetAdviceFeeKind(lngAdviceID As Long) As Byte
      '���ܣ�����ҽ��ID��ȡ�������͵ķ��õ��ݵ����ʣ�1=������ã�2=סԺ����
          Dim rsTmp As ADODB.Recordset, strSQL As String

1         On Error GoTo GetAdviceFeeKind_Error


2         GetAdviceFeeKind = 2
3         strSQL = "Select a.��¼����,a.�������,b.������Դ From ����ҽ������ a,����ҽ����¼ b Where a.ҽ��ID = [1] and a.ҽ��id= b.id"

4         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��ѯ�걾ID", lngAdviceID)
5         If rsTmp.RecordCount > 0 Then
6             If rsTmp!��¼���� = 1 Or rsTmp!��¼���� = 2 And Val("" & rsTmp!�������) = 1 Then
7                 GetAdviceFeeKind = 1
8             Else
9                 If Val("" & rsTmp!������Դ) = 4 Then
10                    GetAdviceFeeKind = 1
11                End If
12            End If
13        End If

14        Exit Function



15        Exit Function
GetAdviceFeeKind_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetAdviceFeeKind)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear

End Function

Public Function funGetNewDataToXK(ByVal lngPatientID As Long, ByVal strItemCode As String) As String
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                  ���ݴ��벡��ID��ָ����� ���ص�ǰ�������һ�εļ�����
          '����
          '                       lngPatientID ����id
          '                       strItemCode ָ����봮��ʹ�� ���ŷָ�
          '                      ���ظ�ʽ��   ����id<A>ָ�����1<S>������1<B>ָ�����2<S>������2<B>ָ�����3<S>������3

          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTest As ADODB.Recordset
          Dim strErr As String, intloop As Integer
          Dim varItemCode As Variant
          Dim strData As String

1         On Error GoTo funGetNewDataToXK_Error

2         varItemCode = Split(strItemCode, ",")
3         For intloop = 0 To UBound(varItemCode)
4             strSQL = "Select *" & vbNewLine & _
                       "   From (Select d.ָ�����, c.������, b.���ʱ��" & vbNewLine & _
                       "          From ���鱨���¼ B, ���鱨����ϸ C, ����ָ�� D" & vbNewLine & _
                       "          Where b.Id = c.�걾id And c.��Ŀid = d.Id And b.His����id = [1]  And d.ָ����� = [2] And b.����� Is Not Null " & vbNewLine & _
                       "          Order By b.���ʱ�� Desc)" & vbNewLine & _
                       "   Where Rownum < 2"
5             Set rsTest = ComOpenSQL(Sel_Lis_DB, strSQL, "��ѯ������", lngPatientID, varItemCode(intloop))
6             If rsTest.RecordCount > 0 Then
7                 strData = strData & "<B>" & rsTest("ָ�����") & "<S>" & rsTest("������")
8             End If
              
9         Next
10        If strData <> "" Then strData = Mid(strData, 4)
11        funGetNewDataToXK = lngPatientID & "<A>" & strData

12        Exit Function
funGetNewDataToXK_Error:
13        strErr = "funGetNewData����" & Err.Number & " " & Err.Description
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetNewDataToXK)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
15        Err.Clear
End Function

Public Function funGetReadNotify(objFrm As Object, strAdvice As String, ByVal strDicName As String, Optional ByRef strReturn As String) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                   ҽ������վ����Σ��ֵ�����ʩ��LIS
          '����
          '       ���
          '                       objfrm          ���ô���
          '                       strAdvice       ҽ��id��
          '                       strDicName      ҽ������
          '       ����
          '                       strReturn       ����ҽ����д�Ĵ����ʩ
          '����                   True=�ɹ�,False=ʧ��
          '
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

          Dim lngSampleID As Long, strTime As String

1         On Error GoTo funGetReadNotify_Error

2         If strAdvice <> "" Then
3             strSQL = "select b.id,b.����,b.�걾��� from ����������� a,���鱨���¼ b " & vbNewLine & _
                          "where a.�걾id = b.id  and a.����id =[1]"

4             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "Σ��ֵ��ѯ", CLng(strAdvice))
5             If rsTmp.EOF = False Then
6                 strTime = Format(Currentdate, "yyyy-mm-dd hh:mm:ss")
7                 lngSampleID = Val(rsTmp("id") & "")
8                 funGetReadNotify = frmAppforCritical.ShowMe(objFrm, strDicName, lngSampleID, strReturn)
9             End If
              
10        End If
          


11        Exit Function
funGetReadNotify_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetReadNotify)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear
          
End Function

Public Function funGetLISPatientRecord(ByVal lngPatientID As Long, ByVal strRecordID As String, ByVal intType As Integer, Optional strErr As String) As ADODB.Recordset
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                       ��ѯLIS������Ϣ
          '
          '���
          '                           lngPatientID    ����ID
          '                           lngRecordID     intType=1 �Һŵ� intType=2 ��ҳID intType=3 ���
          '                           intType         1=���� 2=סԺ 3=����
          '                           strErr          ���صĴ�����Ϣ
          '����
          '                           ������Ϣ��¼��,��¼���а������²�����Ϣ
          '                           ҽ��ID,���ID,����ʱ��,������Դ,Ӥ��,����,�Ա�,����,��������,���䵥λ,������, �������,����,���˿���,����,�Һŵ�,�����,סԺ��,��ҳID,�걾����,���� from �����������
          '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo funGetLISPatientRecord_Error

2         strSQL = "select ����ID as ҽ��ID,ҽ��ID as ���ID,����ʱ��,������Դ,Ӥ��,����,�Ա�,����,��������,���䵥λ,������," & _
                 " �������,����,���˿���,����,�Һŵ�,�����,סԺ��,��ҳID,�걾����,���� from ����������� where HIS����ID=[1]"
          
3         Select Case intType
              Case 1
4                 strSQL = strSQL & " and �Һŵ�=[2] and ������Դ=1"
5             Case 2
6                 strSQL = strSQL & " and ��ҳID=[2] and ������Դ=2"
7             Case 3
8                 strSQL = strSQL & " and �Һŵ�=[2] and ������Դ=4"
9         End Select
          
10        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ѯLIS����", lngPatientID, strRecordID)
11        Set funGetLISPatientRecord = rsTmp

12        Exit Function
funGetLISPatientRecord_Error:
13        strErr = Err.Number & "   " & Err.Description
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetLISPatientRecord)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
15        Err.Clear
End Function

Public Function funModifyPatientBaseIntoLIS(ByVal lng����ID As Long, ByVal str����ID As String, ByVal int���� As Integer, ByVal strName As String, _
                                        ByVal strSex As String, ByVal lngAgeNum As Long, ByVal strAgeUnit As String, ByVal strEditMode As String, _
                                        ByVal strEditUser As String, Optional strErr As String) As Boolean

          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                       ͬ���޸ĵ�ZLLIS������Ϣ��
          '
          '���
          '                           lng����ID
          '                           str����ID       ����=1 �Һŵ� ����=2 ��ҳID ����=3 ���
          '                           int����         1=���� 2=סԺ 3=����
          '                           strName         Ҫ�޸ĵĲ�������
          '                           strSex          Ҫ�޸ĵĲ����Ա�
          '                           strAgeNum       Ҫ�޸ĵĲ�����������
          '                           strAgeUnit      Ҫ�޸ĵĲ������䵥λ
          '
          '                           strEditMode     �޸�Դ���ĸ�ģ��
          '                           strEditUser     �޸���
          '                           strErr          ���صĴ�����Ϣ
          '����
          '                           True=����ɹ� False=����ʧ��
          '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strAgeAll As String    '����
          Dim intBaby As Integer  'Ӥ��
          
          Dim strAgeUnit1 As String    '��һ���䵥λ
          Dim strAgeUnit2 As String   '�ڶ����䵥λ
          Dim strAge1 As String        '��һ����
          Dim strAge2 As String       '�ڶ�����
          Dim strInfo As String

          
1         On Error GoTo funModifyPatientBaseIntoLIS_Error

2         funModifyPatientBaseIntoLIS = False

3         Set rsTmp = funGetLISPatientRecord(lng����ID, str����ID, int����, strErr)
          
          'û�в��ҵ���Ϣ���˳�
4         If rsTmp.RecordCount < 1 Then
5             funModifyPatientBaseIntoLIS = True
6             Exit Function
7         Else
8             intBaby = Val(rsTmp("Ӥ��") & "")
9         End If
          
          '�Ա�
10        Select Case strSex
              Case "��", 1
11                strSex = "1"
12            Case "Ů", 2
13                strSex = "2"
14            Case "δ֪", 9
15                strSex = "9"
16            Case Else
17                strSex = "0"
18        End Select
          
          '����
19        strAgeAll = lngAgeNum & strAgeUnit
20        If InStr(strAgeAll, "��") > 0 Then
21            strAgeUnit1 = "��"
22            strAgeUnit2 = "��"
23        ElseIf InStr(strAgeAll, "��") > 0 Then
24            strAgeUnit1 = "��"
25            strAgeUnit2 = "��"
26        ElseIf InStr(strAgeAll, "��") > 0 Then
27            strAgeUnit1 = "��"
28            strAgeUnit2 = "ʱ"
29        ElseIf InStr(strAgeAll, "ʱ") > 0 Then
30            strAgeUnit1 = "ʱ"
31            strAgeUnit2 = "��"
32        ElseIf InStr(strAgeAll, "��") > 0 Then
33            strAgeUnit1 = "��"
34            strAgeUnit2 = ""
35        ElseIf InStr(strAgeAll, "Ӥ") > 0 Then
36            strAgeUnit1 = "Ӥ"
37            strAgeUnit2 = ""
38        ElseIf InStr(strAgeAll, "��") > 0 Then
39            strAgeUnit1 = "��"
40            strAgeUnit2 = ""
41        End If
42        strAgeAll = Replace(strAgeAll, "����", "��")
43        strAgeAll = Replace(strAgeAll, "Сʱ", "ʱ")
44        strAgeAll = Replace(strAgeAll, "����", "��")
          
45        strAge1 = Mid(strAgeAll, 1, InStr(strAgeAll, strAgeUnit1) - 1)
46        strAge2 = Mid(strAgeAll, InStr(strAgeAll, strAgeUnit1) + 1)
47        strAge2 = Replace(strAge2, strAgeUnit2, "")
48        strInfo = CheckAgeInfo(strAge1, strAgeUnit1, strAge2, strAgeUnit2, strAgeAll)
49        If strInfo <> "" Then
50            strErr = strInfo
51            funModifyPatientBaseIntoLIS = False
52            Exit Function
53        End If
            
          
54        strSQL = "Zl_Lis������Ϣ_����(" & lng����ID & ",'" & str����ID & "'," & int���� & ",'" & strName & "','" & strSex & _
                 "','" & strAgeAll & "'," & lngAgeNum & ",'" & strAgeUnit & "','" & strEditMode & "','" & strEditUser & "'," & intBaby & ")"
55        Call ComExecuteProc(Sel_Lis_DB, strSQL, "�޸Ĳ�����Ϣ")
          
56        strErr = "LIS������Ϣ���޸�,��֪ͨ�����������˲��˱���"
57        funModifyPatientBaseIntoLIS = True



58        Exit Function
funModifyPatientBaseIntoLIS_Error:
59        strErr = Err.Description
60        If InStr(strErr, "[ZLSOFT]") > 0 Then
61            strErr = Mid(strErr, InStr(strErr, "[ZLSOFT]") + 8, InStrRev(strErr, "[") - InStr(strErr, "[ZLSOFT]") - 8)
62        End If
63        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funModifyPatientBaseIntoLIS)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
64        Err.Clear
End Function

Public Function CheckAgeInfo(strAge As String, strAgeUnit As String, strAge1 As String, strAgeUnit1 As String, Optional ByVal strFullAge As String) As String
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����               ���������Ƿ�ϸ�Ҫ��
          '����
          '                   strAge  = ���䣨2��
          '                   strAgeUnit = ���䵥λ (��)
          '                   strAge1 = �ڶ����䣨3��
          '                   strAgeUnit1 = �ڶ����䵥λ(�£�
          '                   strfullAge  =�����������ַ���
          '����               ��ȷ����Ϊ���ִ������󷵻������������ʾ����
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strTmp As String

          '�ж������ַ�����һ���ַ��Ƿ�Ϊ����
1         On Error GoTo CheckAgeInfo_Error

2         If InStr("0123456789", Mid(strFullAge, 1, 1)) <= 0 And strFullAge <> "" Then
3             CheckAgeInfo = "���䲻����Ҫ��,�������ֲ���Ϊ�����֣�"
4             Exit Function
5         End If
          
          '�жϵ�һ����
6         If IsNumeric(strAge) = False And Val(Trim(strAge)) <> 0 Then
7             CheckAgeInfo = "���䲻����Ҫ�󣬲���ȫ���֣�"
8             Exit Function
9         End If
          '�ж������С
10        If Val(strAge) > 150 And strAgeUnit = "��" Then
11            CheckAgeInfo = "���䲻�ܳ���150�꣡"
12            Exit Function
13        End If

          '�ж����䵥λ
14        If Trim(strAgeUnit) = "" And Val(Trim(strAge)) <> 0 Then
              '������������ʱ�����䵥λ����Ϊ��
15            CheckAgeInfo = "������������ʱ�����䵥λ����Ϊ�գ�"
16            Exit Function
17        End If

18        If InStr(",��,��,��,ʱ,��,Ӥ,��,", "," & strAgeUnit & ",") <= 0 And Val(Trim(strAge)) <> 0 Then
19            CheckAgeInfo = "���䵥λ������Ҫ�����飡"
20            Exit Function
21        End If

         '�жϵڶ������Ƿ�Ϊ����
22        If InStr("0123456789", Mid(strFullAge, Len(strAge) + Len(strAgeUnit) + 1, 1)) <= 0 Then
23             CheckAgeInfo = "�ڶ���������֣����飡"
24             Exit Function
25        End If
          
26        If IsNumeric(strAge1) = False And Val(Trim(strAge1)) <> 0 Then
27            CheckAgeInfo = "���䲻����Ҫ�󣬲���ȫ���֣�"
28            Exit Function
29        End If
         
          
30        If Trim(strAgeUnit) <> "" Then
31            Select Case strAgeUnit
                  Case "��"
32                    strTmp = "��"
33                Case "��"
34                    strTmp = "��"
35                Case "��"
36                    strTmp = "ʱ"
37                Case "ʱ"
38                    strTmp = "��"
39                Case Else
40                    strTmp = ""
41            End Select
42            If strTmp <> strAgeUnit1 And strAgeUnit1 <> "" Then
43                CheckAgeInfo = "�ڶ����䵥λ���������飡"
44                Exit Function
45            End If
46        End If
          
47        CheckAgeInfo = ""


48        Exit Function
CheckAgeInfo_Error:
49        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(CheckAgeInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
50        Err.Clear
End Function



Public Function funModifyBabyInfo(ByVal lngBID As Long, ByVal lngZID As Long, ByVal strNO As String, ByVal lngBabyID As Long, ByVal strBabyName As String, ByVal strBabySex As String, Optional strErr As String) As Boolean
          '����        �޸�������������Ϣ
          'lngBID      ����id
          'lngZID      ��ҳid
          'strNO       �Һŵ�
          'lngBabyID   Ӥ�����
          'strBabyName Ӥ������
          'strBabyAge  Ӥ���Ա�
          '����        True=�޸ĳɹ�,Flase=�޸�ʧ��
          
          Dim strSQL As String

1         On Error GoTo funModifyBabyInfo_Error

2         funModifyBabyInfo = False
          
3         strSQL = "Zl_Lis��������Ϣ_Update(" & lngBID & "," & lngZID & ",'" & strNO & "'," & lngBabyID & ",'" & strBabyName & "','" & strBabySex & "')"
4         Call ComExecuteProc(Sel_Lis_DB, strSQL, "�޸���������Ϣ")
          
5         funModifyBabyInfo = True
          

6         Exit Function
funModifyBabyInfo_Error:
7         strErr = "������(funModifyBabyInfo),������Ϣ:" & Err.Number & " " & Err.Description
8         funModifyBabyInfo = False
9         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funModifyBabyInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
10        Err.Clear
End Function

'---------------------------------------------------------------------------------------
' ��    ��:������
' ��������:2017-3-23
' ��    ��:д�����֪ͨ��¼
' ��    ��:
'           intType                 1=����ƾ��ձ걾��2=�ٴ���Ⱦ��������
'           strAdviceID             ����ҽ��ID�����ʹ��","�ָ�
'           [��ѡ][intSampleType    �걾״̬    �ɼ�����վ����ʱ���� 1=�ò����飬2=ȡ�����飬3=�����ز�]
' ��    ��:
'          strErr                  ������Ϣ
' ��    ��:True=�ɹ���False=ʧ��
' �� �� ��:
' �޸�����:
'---------------------------------------------------------------------------------------
Public Function funWriteInLisNotify(ByVal intType As Integer, ByVal strAdviceID As String, _
                                    Optional ByVal intSampleType As Integer, Optional strErr As String) As Boolean
    Dim strSQL As String
    Dim var_tmp As Variant
    Dim strNotify As String
    Dim strBusiness As String
    Dim lngLoop As Long
        
    '��������ҽ����һλΪ���ţ����ȡ����һλ
    On Error GoTo funWriteInLisNotify_Error

    If Mid(strAdviceID, 1, 1) = "," Then
        strAdviceID = Mid(strAdviceID, 2)
    End If
    If strAdviceID = "" Then Exit Function
    
    If intType = 1 Then
        If intSampleType = 1 Then
            strNotify = "��ִ���ò�����"
        ElseIf intSampleType = 2 Then
            strNotify = "ȡ������"
        ElseIf intSampleType = 3 Then
            strNotify = "�������ز�"
        End If
        strBusiness = "�걾����"
    ElseIf intType = 2 Then
        strNotify = "ҽ���ѷ�����Ⱦ���������"
        strBusiness = "��Ⱦ��"
    End If
    
    var_tmp = Split(strAdviceID, ",")
    For lngLoop = LBound(var_tmp) To UBound(var_tmp)
        strSQL = "Zl_������Ϣ��¼_Edit(1,3, Null, Null," & var_tmp(lngLoop) & ",Null,Null,'" & strNotify & "','" & strBusiness & "')"
        Call ComExecuteProc(Sel_Lis_DB, strSQL, "������Ϣ��¼")
    Next
    funWriteInLisNotify = True
    
    Exit Function
funWriteInLisNotify_Error:
    strErr = Err.Description
    Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funWriteInLisNotify)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
    Err.Clear
    
End Function

Public Function PrintReportNew(objFrm As Object, lngAdive As Long, lngPaint As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
          '����       ��ӡ����
          Dim intCount As Integer
          Dim strNO As String
          Dim intSel As Integer
          Dim strChart(0 To 8) As String
          Dim strSQL As String
          Dim strTmp As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset
          Dim lngSampleID As Long
          
1         On Error GoTo PrintReportNew_Error

2         strSQL = "select  b.id from ����������� a ,���鱨���¼  b  where  a.�걾id = b.id and   a.ҽ��id = [1] and a.his����id= [2]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngAdive, lngPaint)
4         If rsTmp.RecordCount > 0 Then
5             lngSampleID = Val(rsTmp("ID") & "")
6         Else
7             strSQL = "select  b.id from ����������� a ,���鱨���¼  b  where  a.�걾id = b.id and   a.����id = [1] and a.his����id= [2]"
8             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngAdive, lngPaint)
9             If rsTmp.RecordCount > 0 Then
10                lngSampleID = Val(rsTmp("ID") & "")
11            Else
12                PrintReportNew = False
13                Exit Function
14            End If
15        End If

16        strSQL = "select b.id ����id ,b.���� ��������,b.�������,a.������Դ,a.����ʱ��,a.���Ա���,a.�걾��� from ���鱨���¼ a,����������¼ b where a.����id = b.id and a.id = [1]"
17        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngSampleID)

18        If rsTmp.RecordCount = 0 Then Exit Function

19        strSQL = "select id,����,����,���ﵥ��,סԺ����,��쵥��,Ժ�ⵥ��,�����ʽ,סԺ��ʽ,����ʽ,Ժ���ʽ,��ʽ����," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(���ﵥ��, '00000')) || '-2' ���ﵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ����, '00000')) || '-2' סԺ���ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(��쵥��, '00000')) || '-2' ��쵥�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ�ⵥ��, '00000')) || '-2' Ժ�ⵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(�����ʽ, '00000')) || '-2' �����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ��ʽ, '00000')) || '-2' סԺ��ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(����ʽ, '00000')) || '-2' ����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ���ʽ, '00000')) || '-2' Ժ���ʽ��" & vbNewLine & _
                      "from ����������¼ where id = [1] "

20        Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", Val(rsTmp("����ID") & ""))


21        rsReportFormat.Filter = "id=" & Val(rsTmp("����ID") & "")
22        If Val(rsTmp("�������")) = 1 Then
23            If Val(rsTmp("���Ա���") & "") = 1 Then
                  '����
24                intSel = 0
25            Else
                  '����
26                intSel = 1
27            End If
28        Else
29            intCount = GetSampleValCount(lngSampleID)
              'û�н��ʱ��ʾ
30            If intCount = 0 Then
31                Exit Function
32            End If
33            If rsReportFormat.RecordCount > 0 Then
34                If Val(rsReportFormat("��ʽ����") & "") > 0 Then
35                    If intCount > Val(rsReportFormat("��ʽ����") & "") Then
36                        intSel = 0
37                    Else
38                        intSel = 1
39                    End If
40                End If
41            Else
42                intSel = 0
43            End If

44        End If
45        Select Case Val(rsTmp("������Դ") & "")
              Case 1
46                If intSel = 0 Then
47                    strNO = rsReportFormat("���ﵥ�ݺ�")
48                Else
49                    strNO = rsReportFormat("�����ʽ��")
50                End If
51            Case 2
52                If intSel = 0 Then
53                    strNO = rsReportFormat("סԺ���ݺ�")
54                Else
55                    strNO = rsReportFormat("סԺ��ʽ��")
56                End If
57            Case 3
58                If intSel = 0 Then
59                    strNO = rsReportFormat("סԺ���ݺ�")
60                Else
61                    strNO = rsReportFormat("סԺ��ʽ��")
62                End If
63            Case 4
64                If intSel = 0 Then
65                    strNO = rsReportFormat("Ժ�ⵥ�ݺ�")
66                Else
67                    strNO = rsReportFormat("Ժ���ʽ��")
68                End If
69            Case Else
70                If intSel = 0 Then
71                    strNO = rsReportFormat("���ﵥ�ݺ�")
72                Else
73                    strNO = rsReportFormat("�����ʽ��")
74                End If
75        End Select
76        If byRunMode = 3 Then
77            If strNO <> "" Then
78                 FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
79            End If
80        Else
             '��ͼ��
81            strTmp = "��ʼ����ͼ��:" & Now & vbCrLf
82            If ReadSampleImage(lngSampleID, strChart, strErr) = False Then
83                MsgBox strErr: Exit Function
84            End If
85            strTmp = strTmp & "����ͼ�����:" & Now & vbCrLf
          
86            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "�걾ID=" & lngSampleID, "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), _
                      "ͼ��4=" & strChart(3), "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                      "ͼ��9=" & strChart(8), byRunMode
87            strTmp = strTmp & "��ӡ���:" & Now & vbCrLf
              
              '������˹��ı걾��ʶ
88            strSQL = "Zl_���鱨���ӡ_Edit(1," & lngSampleID & ")"
89            Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
90            strTmp = strTmp & "��ɴ�ӡ:" & Now
          
91            SaveDBLog 18, 6, lngSampleID, "��ӡ", "�����ӡ", 2500, "�ٴ�ʵ���ҹ���"
92        End If

93        PrintReportNew = True


94        Exit Function
PrintReportNew_Error:
95        PrintReportNew = False
96        strErr = "ִ��(PrintReportNew)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl
97        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(PrintReportNew)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
98        Err.Clear
End Function

Public Function GetSampleValCount(lngSampleID As Long, Optional strErr As String) As Integer
          '����           ȡ��ǰ�걾��������

          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo GetSampleValCount_Error

2         strSQL = "select count(*) count from ���鱨����ϸ where �걾id = [1] and ������ is not null "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", lngSampleID)
4         GetSampleValCount = rsTmp("count")

5         Exit Function
GetSampleValCount_Error:
6         strErr = WriteErrLog("zl9LisInsideComm", "mdlLisHisComm", "ִ��(GetSampleValCount)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear

End Function

Public Function ReadSampleImage(lngSampleID As Long, strChar() As String, Optional strErr As String, Optional intVal As Integer = 25) As Boolean
    '����   ����걾��ͼ�񷵻ض���������
    '��ͼ��
    Dim strReturn As String
    Dim varTmp As Variant, strDir As String
    Dim i As Integer
    Dim gobjFSO As New Scripting.FileSystemObject    'FSO����
    Dim objImg As Object

    On Error GoTo ReadSampleImage_Error

    strErr = ""
    strDir = App.Path & "\LisImage"
    If Not gobjFSO.FolderExists(strDir) Then Call gobjFSO.CreateFolder(strDir)

    If objImg Is Nothing Then
        Set objImg = CreateObject("zlLisDev.clsDrawGraph")



        If strErr <> "" Then
            MsgBox strErr
            Exit Function
        End If
    End If
    objImg.GetSampleImgExit strErr
    '�걾ID
    'ͼƬ����·��(���������Զ�����),
    '�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
    '��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
    '���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
    '���°�LIS�����ϰ�LIS�ڵ��ñ��������� 0-�ϰ�LIS��Ĭ�ϣ��ӡ�����ͼ��������ȡͼ�����ݣ���1-�°�LIS���ӡ����鱨��ͼ����ȡͼ�����ݣ�
    If intVal = 25 Then
         Call objImg.GetSampleImgInit(gSysInfo.SysNo, gcnLisOracle, strErr)
        strReturn = objImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 1)
    Else
         Call objImg.GetSampleImgInit(gSysInfo.SysNo, gcnHisOracle, strErr)
        strReturn = objImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 0)
    End If
    If strReturn = "" Then
        If strErr = "��ͼ�����ݣ�" Then
            strErr = ""
            ReadSampleImage = True
        ElseIf strErr <> "" Then
            MsgBox strErr, vbQuestion
        Else
            ReadSampleImage = True
        End If
        Exit Function
    End If

    varTmp = Split(strReturn, ",")

    For i = LBound(varTmp) To UBound(varTmp)
        If i > 8 Then Exit For
        If Trim("" & varTmp(i)) <> "" Then
            If Dir(strDir & "\" & Trim("" & varTmp(i))) <> "" Then strChar(i) = strDir & "\" & Trim("" & varTmp(i))
        End If
    Next

    ReadSampleImage = True

    Exit Function
ReadSampleImage_Error:
    strErr = "������(ReadSampleImage),������Ϣ:" & Err.Number & " " & Err.Description
    Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(ReadSampleImage)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
    Err.Clear
End Function



'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/9
'��    ��:��lis��Ϣ��������ˢ���ʿظøſ��б���Ϣ
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Sub SendMessage(ByVal strMessage As String)
1         On Error GoTo SendMessage_Error
          
2         If mstrPara = "" Then mstrPara = ComGetPara(Sel_Lis_DB, "LISԶ��ͨѶ����", 2500, 2500, "")
3         If mstrPara = "" Then Exit Sub
4         If gobjPublicLIS Is Nothing Then
5             Set gobjPublicLIS = CreateObject("zlPublicLIS.clsSampleReprot")
6             If Not gobjPublicLIS Is Nothing Then Call gobjPublicLIS.Init(mstrPara)
7         End If
          
8         If Not gobjPublicLIS Is Nothing Then
9             Call gobjPublicLIS.SendMessage(strMessage, mstrPara)
10        End If


11        Exit Sub
SendMessage_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(SendMessage)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear

End Sub







'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/8/25
'��    ��:  ��ѯ����סԺ�ڼ������ѳ����������ID������ͨ���걾���з���
'��    ��:
'           lngPatientID    HIS����ID
'           intPage         ��ҳID
'��    ��:
'           strErr          ������Ϣ
'��    ��:  ���ذ��ձ걾���з����ҽ��ID����ҽ��֮����","�ָ�걾֮����";"�ָ�
'---------------------------------------------------------------------------------------
Public Function funGetPatientAdvice(ByVal lngPatientID As Long, ByVal intPage As Integer, Optional ByRef strErr As String) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strAdvice As String
          
1         On Error GoTo funGetPatientAdvice_Error
          
2         strErr = ""
          
3         strSQL = "select f_List2str(Cast(Collect(a.����ID || '') As t_Strlist)) ����ID" & vbCrLf & _
                  " from  ����������� A,���鱨���¼ B " & vbCrLf & _
                  " where a.�걾id=b.id and a.his����ID=[1] and a.��ҳID=[2] and b.����� is not null group by a.�걾ID"
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", lngPatientID, intPage)
5         If rsTmp.RecordCount <= 0 Then Exit Function
6         Do While Not rsTmp.EOF
7             strAdvice = strAdvice & ";" & rsTmp("����ID")
8             rsTmp.MoveNext
9         Loop
10        If Mid(strAdvice, 1, 1) = ";" Then strAdvice = Mid(strAdvice, 2)
11        funGetPatientAdvice = strAdvice

12        Exit Function
funGetPatientAdvice_Error:
13        strErr = Err.Description & "(" & Err.Number & ")"
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetPatientAdvice)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
15        Err.Clear
End Function




'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/11/22
'��    ��:��дΣ��ֵ�����ʩ��LISϵͳ��
'��    ��:
'           lngSampleID     �걾ID
'           strUserName     ������Ա
'           strNotify       �����ʩ
'��    ��:
'           [strErr         ������Ϣ]
'��    ��:  ��д�ɹ�����True,���򷵻�False
'---------------------------------------------------------------------------------------
Public Function funWriteNotifyToLis(ByVal lngSampleID As Long, ByVal strUserName As String, ByVal strNotify As String, Optional ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim strTime As String
          Dim blnTrs As Boolean
          
1         On Error GoTo funWriteNotifyToLis_Error
          
2         strTime = Currentdate
          
3         gcnLisOracle.BeginTrans
4         blnTrs = True
           '������Ϣ����
5         strSQL = "Zl_������Ϣ��¼_Edit(1,2,null," & lngSampleID & ",null,null,null,'ҽ���Ѳ���Σ��ֵ','Σ��ֵ')"
6         Call ComExecuteProc(Sel_Lis_DB, strSQL, "������Ϣ��¼")
          
7         strSQL = "Zl_����Σ��ֵ��¼_Message(" & lngSampleID & ",'" & strUserName & "',to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),'" & strNotify & "')"
8         Call ComExecuteProc(Sel_Lis_DB, strSQL, "Σ��ֵ��¼֪ͨ")
9         gcnLisOracle.CommitTrans
10        blnTrs = False
          
11        SaveDBLog 18, 6, Val(lngSampleID), "Σ��ֵ����", "ȷ�ϴ���ȷ���ˣ�" & strUserName & " ȷ��ʱ��:" & Format(strTime, "yyyy-MM-dd HH:mm:ss") & " �����ʩ:" & strNotify, 2500, "�ٴ�ʵ���ҹ���"
          
12        funWriteNotifyToLis = True

13        Exit Function
funWriteNotifyToLis_Error:
14        strErr = Err.Description
15        If blnTrs Then gcnLisOracle.RollbackTrans
16        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funWriteNotifyToLis)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
17        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/14
'��    ��:  �ж��Ƿ�������������﷢�͵�ҽ��
'��    ��:
'           strAdvicIDs     ҽ��ID�����ҽ��ʹ�á�,���ָ�
'��    ��:
'��    ��:  True=�����ﲡ�ˣ�False=�������ﲡ��
'---------------------------------------------------------------------------------------
Public Function funNewSystemSvr(ByVal strAdvicIDs As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strGHNo As String
          Dim astrtmp() As String
          Dim strJsOut As String          '���ص�JSON
          Dim intFinish As Integer
          Dim i As Integer
          Dim strErr As String
          Dim blnTmp As Boolean

          '    json˵��
          '    �ֶ�       ����        ˵��
          '    result     ִ�н��    1-�ɹ���-1-ʧ�� Number(1)   �ǿ�
          '    errmsg     ������Ϣ    ʧ��ʱ���ش�����Ϣ  Varchar2(200)
          '    kacnt_sign �շ�״̬    0-δ�շѣ�1-���շѣ��ñ�־��ʾִ�п����Ƿ��ִ�У���ɫͨ������1���ռ䲡��Ԥ�����㹻����1���˵����շѷ���1����������0    Number(1)   �ǿ�
          '    kacnt_chrg δ�ս��    ��ɫͨ������Ԥ�����㲿�ֽ���������δ�շѽ��  Number(18,2)    �ǿ�


          '��ѯסԺ��¼
1         On Error GoTo funNewSystemSvr_Error

2         strSQL = "select /*+cardinality(c,10)*/  b.���ӱ�־" & vbCrLf & _
                   " from ����ҽ����¼ A,���˹Һż�¼ B,Table(f_Num2list([1])) C " & vbCrLf & _
                   " where a.�Һŵ�=b.no and a.id=c.Column_Value" & vbCrLf & _
                   "union all " & vbCrLf & _
                   "select /*+cardinality(c,10)*/ b.���ӱ�־" & vbCrLf & _
                   " from ����ҽ����¼ A,���˹Һż�¼ B,Table(f_Num2list([1])) C " & vbCrLf & _
                   " where a.�Һŵ�=b.no and a.���id=c.Column_Value"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ҳ", strAdvicIDs)
4         Do While Not rsTmp.EOF
5             If Val(rsTmp("���ӱ�־") & "") = 3 Then    '3��ʾ�����ﲡ��

                  '���������ﲡ����������ɼ�վ�Ѿ��ϸ�������շ�״̬������LISϵͳ�в��ڼ�飬�����ﲡ��һ�ɲ�����
6                 funNewSystemSvr = True
7                 Exit Function

8             End If
9             rsTmp.MoveNext
10        Loop

11        Exit Function
funNewSystemSvr_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funNewSystemSvr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear

End Function

Public Function funGetSampleType(ByVal strAdvice As String, Optional strErr As String) As ADODB.Recordset
      '����       ��ȡ�걾״̬

      '����
      '           strAdvice       ҽ��ID,���ҽ����","�ָ�
      '           strErr          ���ش�����Ϣ

      '���ؼ�¼��
      '��¼���ֶ�: "ҽ��ID", adBigInt
      '           "ҽ��״̬", adVarChar, 20
      '           "����Ա", adVarChar, 20
      '           "����ʱ��", adDate

          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset    'ҽ������
          Dim strType As String           '�걾״̬
          Dim strReturn As String         '���ؽ��
1         Dim strAdviceID As String       'ҽ��ID��      ��ʽ:�ϰ�ҽ��ID1,�ϰ�ҽ��ID2,,|�°�ҽ��ID1,�°�ҽ��ID2,,,
          Dim strOldAdvice As String      '�ϰ�ҽ��ID
          Dim strNewAdvice As String      '�°�ҽ��ID
          Dim rsReture As ADODB.Recordset    '���صļ�¼��
          Dim strUser As String           '����Ա
          Dim strDate As String           '����ʱ��
          Dim intType As Integer
          Dim var_tmp As Variant
          Dim intloop As Integer
          Dim strArr As Variant
          Dim i As Integer

          '���Ի����ؼ�¼��
2         On Error GoTo funGetSampleType_Error

3         Set rsReture = InitRecord

4         strArr = TruncatedExtraLongStr(strAdvice, ",")

5         For i = 0 To UBound(strArr)
6             strOldAdvice = ""
7             strNewAdvice = ""
8             strAdviceID = ""
              
              '��������֮ǰ����
9             strSQL = "select /*+cardinality(c,10)*/ distinct a.ҽ��ID ����ID ,e.ִ��״̬,a.������,a.����ʱ��,a.�ͼ���," & _
                     " a.�걾�ͳ�ʱ�� �ͼ�ʱ��,a.������,a.����ʱ��,b.������,b.����ʱ��,'' ������,'' ����ʱ��," & _
                     " b.�����,b.���ʱ�� from ����ҽ������ A,����걾��¼ B,Table(f_num2list([1])) C,����ҽ����¼ D,����ҽ������ E " & _
                     " where a.ҽ��id=b.ҽ��id(+) and a.ҽ��id=d.���id and d.id=e.ҽ��id And a.ҽ��ID=c.Column_Value"
10            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�ϰ�ҽ��", strArr(i))

              '��ȡҽ��״̬
11            Do While Not rsTmp.EOF
12                strType = ""
13                strUser = ""
14                strDate = ""
15                If IsNull(rsTmp("����ʱ��")) And Val(rsTmp("ִ��״̬") & "") <> 2 Then  'δ����
16                    strType = "δ����"
17                ElseIf Val(rsTmp("ִ��״̬") & "") = 2 Then     '�Ѿ���
18                    strType = "�Ѿ���"
19                    strUser = rsTmp("������") & ""
20                    strDate = rsTmp("����ʱ��") & ""
21                End If

22                If Val(rsTmp("ִ��״̬") & "") <> 2 Then
23                    If Not IsNull(rsTmp("����ʱ��")) Then    '�Ѳ���
24                        strType = "�Ѳ���"
25                        strUser = rsTmp("������") & ""
26                        strDate = rsTmp("����ʱ��") & ""
27                    End If
28                    If Not IsNull(rsTmp("�ͼ�ʱ��")) Then    '���ͼ�
29                        strType = "���ͼ�"
30                        strUser = rsTmp("�ͼ���") & ""
31                        strDate = rsTmp("�ͼ�ʱ��") & ""
32                    End If
33                    If Not IsNull(rsTmp("������")) Then   '�ѽ���
34                        strType = "�ѽ���"
35                        strUser = rsTmp("������") & ""
36                        strDate = rsTmp("����ʱ��") & ""
37                    End If
38                End If

                  '��ӵ����ؼ�¼��
39                rsReture.AddNew
40                rsReture("ҽ��ID") = CLng(rsTmp("����ID") & "")
41                rsReture("ҽ��״̬") = strType
42                If strUser <> "" Then
43                    rsReture("����Ա") = strUser
44                End If
45                If strDate <> "" Then
46                    rsReture("����ʱ��") = CDate(Format(strDate, "yyyy-mm-dd hh:mm:ss"))
47                End If
48                rsTmp.MoveNext
49            Loop


              '��������֮������
              '��ѯ�ϰ�ҽ��ID
50            strSQL = "select /*+cardinality(b,10)*/ distinct ҽ��ID from ������Ŀ�ֲ� A,Table(f_num2list([1])) B where a.ҽ��id=b.column_value"
51            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�ϰ�ҽ��", strArr(i))
52            Do While Not rsTmp.EOF
53                strOldAdvice = strOldAdvice & "," & rsTmp("ҽ��ID")
54                rsTmp.MoveNext
55            Loop
56            If strOldAdvice <> "" Then strOldAdvice = Mid(strOldAdvice, 2)

              '��ѯ�°�ҽ��ID
57            strSQL = " select /*+cardinality(b,10)*/ distinct ����ID from ����������� A,Table(f_num2list([1])) B where a.����id=b.column_value"
58            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�°�ҽ��", strArr(i))
59            Do While Not rsTmp.EOF
60                strNewAdvice = strNewAdvice & "," & rsTmp("����ID")
61                rsTmp.MoveNext
62            Loop
63            If strNewAdvice <> "" Then strNewAdvice = Mid(strNewAdvice, 2)

64            strAdviceID = strOldAdvice & "|" & strNewAdvice

65            If strAdviceID <> "" Then var_tmp = Split(strAdviceID, "|")
66            For intloop = LBound(var_tmp) To UBound(var_tmp)
67                If intloop = 0 And var_tmp(0) <> "" Then
                      '��ѯ�ϰ�ҽ����Ϣ
68                    intType = 10
69                    strSQL = "select /*+cardinality(c,10)*/ distinct a.ҽ��ID ����ID ,e.ִ��״̬,a.������,a.����ʱ��,a.�ͼ���," & _
                             " a.�걾�ͳ�ʱ�� �ͼ�ʱ��,a.������,a.����ʱ��,b.������,b.����ʱ��,'' ������,'' ����ʱ��," & _
                             " b.�����,b.���ʱ�� from ����ҽ������ A,����걾��¼ B,Table(f_num2list([1])) C,����ҽ����¼ D,����ҽ������ E " & _
                             " where a.ҽ��id=b.ҽ��id(+) and a.ҽ��id=d.���id and d.id=e.ҽ��id And a.ҽ��ID=c.Column_Value"

70                    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�ϰ�ҽ��", var_tmp(intloop))
71                ElseIf intloop = 1 And var_tmp(1) <> "" Then
                      '��ѯ�°�ҽ����Ϣ
72                    intType = 25
73                    strSQL = "Select /*+cardinality(c,10)*/ distinct a.����ID,'0' ִ��״̬,a.�걾ID ,a.����ʱ��, a.����ʱ��, a.������, a.�ͼ�ʱ��," & _
                             " a.�ͼ���, a.����ʱ��, a.������, b.����ʱ��, b.������ ������, b.������," & _
                             " b.����ʱ��, b.�����, b.���ʱ�� ,a.������,a.����ʱ��,a.�����, a.סԺ��, a.��������" & _
                             " from ����������� A, ���鱨���¼ B,Table(f_num2list([1])) C" & _
                               "��Where a.�걾id = b.Id(+) And a.����id=c.Column_Value"

74                    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�°�ҽ��", var_tmp(intloop))
75                End If


                  '��ȡҽ��״̬
76                Do While Not rsTmp.EOF
77                    strType = ""
78                    strUser = ""
79                    strDate = ""
80                    If Not IsNull(rsTmp("����ʱ��")) Then   '�Ѻ���
81                        If intType = 25 Then
82                            strType = "�Ѻ���"
83                            strUser = rsTmp("������") & ""
84                            strDate = rsTmp("����ʱ��") & ""
85                        ElseIf intType = 10 Then
86                            If Val(rsTmp("ִ��״̬") & "") <> 2 Then
87                                strType = "�Ѻ���"
88                                strUser = rsTmp("������") & ""
89                                strDate = rsTmp("����ʱ��") & ""
90                            End If
91                        End If
92                    End If
93                    If Not IsNull(rsTmp("���ʱ��")) Then   '�����
94                        If intType = 25 Then
95                            strType = "�����"
96                            strUser = rsTmp("�����") & ""
97                            strDate = rsTmp("���ʱ��") & ""
98                        ElseIf intType = 10 Then
99                            If Val(rsTmp("ִ��״̬") & "") <> 2 Then
100                               strType = "�����"
101                               strUser = rsTmp("�����") & ""
102                               strDate = rsTmp("���ʱ��") & ""
103                           End If
104                       End If
105                   End If


                      '���µ�ǰҽ����Ӧ�ļ�����ڲ�״̬
106                   rsReture.Filter = "ҽ��ID=" & CLng(rsTmp("����ID") & "")
107                   If rsReture.RecordCount > 0 Then
108                       If strType <> "" Then
109                           rsReture("ҽ��״̬") = strType
110                       End If
111                       If strUser <> "" Then
112                           rsReture("����Ա") = strUser
113                       End If
114                       If strDate <> "" Then
115                           rsReture("����ʱ��") = CDate(Format(strDate, "yyyy-mm-dd hh:mm:ss"))
116                       End If
117                   End If
118                   rsTmp.MoveNext
119               Loop
120           Next
121       Next

122       rsReture.Filter = ""
123       rsReture.MoveFirst
124       Set funGetSampleType = rsReture


125       Exit Function
funGetSampleType_Error:
126       strErr = "ִ��(BeforCreateLisValueStr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl
127       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetSampleType)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
128       Err.Clear

End Function

Private Function InitRecord(Optional strErr As String) As ADODB.Recordset
          '��ʼ�����ؼ�¼��
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo InitRecord_Error

2         If rsTmp.State = adStateOpen Then rsTmp.Close
3         rsTmp.Fields.Append "ҽ��ID", adBigInt
4         rsTmp.Fields.Append "ҽ��״̬", adVarChar, 20
5         rsTmp.Fields.Append "����Ա", adVarChar, 20
6         rsTmp.Fields.Append "����ʱ��", adDate

7         rsTmp.CursorLocation = adUseClient
8         rsTmp.LockType = adLockOptimistic
9         rsTmp.CursorType = adOpenStatic
10        If rsTmp.State = adStateClosed Then rsTmp.Open

11        Set InitRecord = rsTmp


12        Exit Function
InitRecord_Error:
13        strErr = "ִ��(BeforCreateLisValueStr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(InitRecord)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
15        Err.Clear
          
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/18
'��    ��:��鵱ǰʱ���Ƿ���ҵ��߷��ڣ�����ҵ��ָ���Ĳ�ѯ��Χ�Ƿ�����ɷ�Χ��
'��    ��:
'      lngSysNo=ϵͳ��
'      lngModuleNo=ģ���
'      strFuncName=��������
'      datBegin=���ܽ��в�ѯ���ݷ�Χ�Ŀ�ʼʱ�䣬������Ϊ��ֵ����ʱ
'      datEnd=���ܽ��в�ѯ���ݷ�Χ�Ľ���ʱ��
'      lngDays=��ѯ��ʱ�䷶Χ,��Ϊ0ʱ��ͨ��datBegin��datEnd���㣬��Ϊ0ʱ����datBegin��datEnd
'��    ��:
'��    ��:�Ƿ���Խ��в���
'---------------------------------------------------------------------------------------
Public Function funCheckRushHours(ByVal lngSysNo As Long, ByVal lngModuleNo As Long, ByVal strFuncName As String, _
                                Optional ByVal datBegin As Date, Optional ByVal datEnd As Date, Optional ByVal lngDays As Long) As Boolean
'    If gzlSystem Is Nothing Then
        funCheckRushHours = True
'        Exit Function
'    End If
'    funCheckRushHours = gzlSystem.CheckRushHours(lngSysNO, lngModuleNo, strFuncName, datBegin, datEnd, lngDays)
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/25
'��    ��:��������Ϣд���°�LIS
'��    ��:
'           strAdviceID         ����ID �������IDʹ��","�ָ�
'           strUser             ������
'           strRefuseInfo       ��������
'           strRegName          ���ս�����
'           strRegTime          ���ս���ʱ��

'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Function funRefuseSampleInNew(ByVal strAdviceID As String, ByVal strUser As String, ByVal strRefuseInfo As String, _
                                     Optional ByVal strRegName As String, Optional ByVal strRegTime As String, Optional ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strIDs As String

1         On Error GoTo funRefuseSampleInNew_Error

2         strSQL = "select /*+cardinality(b,10)*/ ID from ����������� where ����ID  In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B)"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", strAdviceID)
4         Do While Not rsTmp.EOF
5             strIDs = strIDs & "," & rsTmp("ID")
6             rsTmp.MoveNext
7         Loop
8         If strIDs <> "" Then strIDs = Mid(strIDs, 2)

9         If VerCompare(gSysInfo.VersionLIS, "10.35.120") = 0 Or VerCompare(gSysInfo.VersionLIS, "10.35.150") <> -1 Then
10            strSQL = "Zl_���鱨�����_Edit('" & strIDs & "','" & strUser & "',null,null,'" & strRefuseInfo & "',0,null," & IIf(strRegName = "", "null", "'" & strRegName & "'") & "," & IIf(strRegTime <> "", "to_date('" & strRegTime & "','yyyy-mm-dd hh24:mi:ss')", "null") & ")"
11        Else
12            strSQL = "Zl_���鱨�����_Edit('" & strIDs & "','" & strUser & "',null,null,'" & strRefuseInfo & "',0,null)"
13        End If
14        Call ComExecuteProc(Sel_Lis_DB, strSQL, "�����������")

15        funRefuseSampleInNew = True


16        Exit Function
funRefuseSampleInNew_Error:
17        strErr = "ִ��(funRefuseSampleInNew)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl
18        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funRefuseSampleInNew)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
19        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/7/5
'��    ��:��ȡ�°�LIS�еı걾����
'��    ��:
'           strInfo     ����ѯ�ظ��ı걾����
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Function funGetSampleTypeNew() As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo funGetSampleTypeNew_Error

2         If intLis_Setup <> 1 Then Exit Function
          
3         strSQL = "select ����,���� from ����걾����"
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�걾����")
5         Set funGetSampleTypeNew = rsTmp


6         Exit Function
funGetSampleTypeNew_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetSampleTypeNew)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
8         Err.Clear
          
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/7/6
'��    ��:ͨ��ҽ��ID��ӡLIS����
'��    ��:
'           objFrm          ���ô���
'           lngAdvice       ҽ��ID���ɼ�ҽ��ID)
'           byRunMode       1=��ӡԤ����2=��ӡ��3=��ӡ���ã�4=��ӡPDF
'           BlnlimitPrint   ��ӡ�°汨��ʱ���Ƿ��ܵ���ӡ�������������ƣ������°�LIS�д�ӡ�����ı����޷���ӡ��
'           strPDF          ��Ҫ��ӡ��PDF�ļ����ļ�·��
'           strPrinter      ָ����ӡ�������ƣ�����ָ���˴�ӡ�����ƣ���Ĭ����ָ���Ĵ�ӡ���ϴ�ӡ
'��    ��:
'           strErr          ���ش�����ӡʧ��ԭ��
'��    ��:  �Ƿ��ӡ�ɹ�    True=�ɹ���False=ʧ��
'---------------------------------------------------------------------------------------
Public Function funPrintLisReport(ByVal objFrm As Object, ByVal lngAdvice As String, ByVal byRunMode As Byte, _
                                  Optional ByVal BlnLimitPrint As Boolean, Optional ByVal strPDF As String, _
                                  Optional ByVal strPrinter As String, Optional ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset
          Dim blnNewReport As Boolean
          Dim blnOldReport As Boolean
          Dim lngPrintCount As Long
          Dim lngSampleID As Long
          Dim lngPaintID As Long
          Dim intSel As Integer
          Dim intCount As Integer
          Dim strNO As String
          Dim strTmp As String
          Dim strChart(0 To 8) As String
          Dim strReportCode As String
          Dim strReportParaNo As String
          Dim bytReportParaMode As Byte
          Dim lngҽ��ID As Long
          Dim lng���ͺ� As Long


1         On Error GoTo funPrintLisReport_Error

          '�ȵ��°�LIS��ȥ��ѯ���鱨��
2         strSQL = "select distinct b.id �걾ID,b.�����,b.ҽ��վ��ӡ,b.������Դ,b.����ID,c.�������,b.���Ա��� " & vbCrLf & _
                 " from ����������� A,���鱨���¼ B,����������¼ C " & vbCrLf & _
                 " where a.�걾ID=b.ID(+) and b.����ID=c.id and ����ID=[1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", lngAdvice)
4         If rsTmp.RecordCount > 0 Then blnNewReport = True

          '���°�LIS��û�б��棬���ٵ��ϰ�LIS��ȥ����
5         If Not blnNewReport Then
6             strSQL = "Select Distinct �걾ID, b.�����, c.���ͺ�, a.ҽ��id, b.����id" & vbCrLf & _
                     "   From ������Ŀ�ֲ� A, ����걾��¼ B, ����ҽ������ C" & vbCrLf & _
                     "   Where a.�걾ID = b.id And a.ҽ��id = c.ҽ��id And a.ҽ��id =[1]"
7             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����걾��¼", lngAdvice)
8             If rsTmp.RecordCount > 0 Then blnOldReport = True
9         End If
          '�жϱ����Ƿ��ѳ���δ���ı����ֹ��ӡ
10        Do While Not rsTmp.EOF
11            If IsNull(rsTmp("�걾ID")) Or IsNull(rsTmp("�����")) Then
12                strErr = "����δ��"
13                Exit Function
14            End If
15            rsTmp.MoveNext
16        Loop

17        If blnNewReport Or blnOldReport Then
18            rsTmp.MoveFirst
19            lngSampleID = Val(rsTmp("�걾ID") & "")
20        End If
          '��ӡ�°汨��ʱ������Ƿ񳬳���ӡ����
21        If BlnLimitPrint = True And blnNewReport = True Then    '��Ҫ��飬�������°�LIS����
22            lngPrintCount = Val(ComGetPara(Sel_Lis_DB, "ҽ������վ�����ӡ����", 2500, 2500, 0))
              '�Աȴ�ӡ�����Ͳ���
23            If lngPrintCount > 0 Then
24                If Val(rsTmp("ҽ��վ��ӡ") & "") >= lngPrintCount And Val(rsTmp("������Դ") & "") = 2 Then
25                    strErr = "������ӡ����"
26                    Exit Function
27                End If
28            End If
29        End If

          '��ӡ����
          '�°�
30        If blnNewReport Then
31            strSQL = "select id,����,����,���ﵥ��,סԺ����,��쵥��,Ժ�ⵥ��,�����ʽ,סԺ��ʽ,����ʽ,Ժ���ʽ,��ʽ����," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(���ﵥ��, '00000')) || '-2' ���ﵥ�ݺ�," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(סԺ����, '00000')) || '-2' סԺ���ݺ�," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(��쵥��, '00000')) || '-2' ��쵥�ݺ�," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(Ժ�ⵥ��, '00000')) || '-2' Ժ�ⵥ�ݺ�," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(�����ʽ, '00000')) || '-2' �����ʽ��," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(סԺ��ʽ, '00000')) || '-2' סԺ��ʽ��," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(����ʽ, '00000')) || '-2' ����ʽ��," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(Ժ���ʽ, '00000')) || '-2' Ժ���ʽ��" & vbNewLine & _
                       "from ����������¼ where id = [1] "

32            Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", Val(rsTmp("����ID") & ""))
33            rsReportFormat.Filter = "id=" & Val(rsTmp("����ID") & "")
34            If Val(rsTmp("�������")) = 1 Then
35                If Val(rsTmp("���Ա���") & "") = 1 Then
                      '����
36                    intSel = 0
37                Else
                      '����
38                    intSel = 1
39                End If
40            Else
41                intCount = GetSampleValCount(lngSampleID)
                  'û�н��ʱ��ʾ
42                If intCount = 0 Then
43                    Exit Function
44                End If
45                If rsReportFormat.RecordCount > 0 Then
46                    If Val(rsReportFormat("��ʽ����") & "") > 0 Then
47                        If intCount > Val(rsReportFormat("��ʽ����") & "") Then
48                            intSel = 0
49                        Else
50                            intSel = 1
51                        End If
52                    End If
53                Else
54                    intSel = 0
55                End If

56            End If
57            Select Case Val(rsTmp("������Դ"))
              Case 1
58                If intSel = 0 Then
59                    strNO = rsReportFormat("���ﵥ�ݺ�")
60                Else
61                    strNO = rsReportFormat("�����ʽ��")
62                End If
63            Case 2
64                If intSel = 0 Then
65                    strNO = rsReportFormat("סԺ���ݺ�")
66                Else
67                    strNO = rsReportFormat("סԺ��ʽ��")
68                End If
69            Case 3
70                If intSel = 0 Then
71                    strNO = rsReportFormat("סԺ���ݺ�")
72                Else
73                    strNO = rsReportFormat("סԺ��ʽ��")
74                End If
75            Case 4
76                If intSel = 0 Then
77                    strNO = rsReportFormat("Ժ�ⵥ�ݺ�")
78                Else
79                    strNO = rsReportFormat("Ժ���ʽ��")
80                End If
81            Case Else
82                If intSel = 0 Then
83                    strNO = rsReportFormat("���ﵥ�ݺ�")
84                Else
85                    strNO = rsReportFormat("�����ʽ��")
86                End If
87            End Select
88            If byRunMode = 3 Then
89                If strNO <> "" Then
90                    FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
91                End If
92            Else
                  '��ͼ��
93                strTmp = "��ʼ����ͼ��:" & Now & vbCrLf
94                If ReadSampleImage(lngSampleID, strChart, strErr, 25) = False Then
95                    Exit Function
96                End If
97                strTmp = strTmp & "����ͼ�����:" & Now & vbCrLf

98                If strPrinter <> "" Then Call FunSetReportPrintSet(gcnLisOracle, gSysInfo.SysNo, strNO, "printer", strPrinter)    '����ָ����ӡ��
99                FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "�걾ID=" & lngSampleID, "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), _
                                "ͼ��4=" & strChart(3), "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                                "ͼ��9=" & strChart(8), "PDF=" & strPDF, byRunMode
100               strTmp = strTmp & "��ӡ���:" & Now & vbCrLf

                  '������˹��ı걾��ʶ
101               strSQL = "Zl_���鱨���ӡ_Edit(1," & lngSampleID & ",1)"
102               Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
103               strTmp = strTmp & "��ɴ�ӡ:" & Now

104               SaveDBLog 18, 6, lngSampleID, "��ӡ", "�������鱨���ӡ", 2500, "�ٴ�ʵ���ҹ���"
105           End If
106       ElseIf blnOldReport Then
107           lng���ͺ� = Val("" & rsTmp("���ͺ�"))
108           lngҽ��ID = Val("" & rsTmp("ҽ��id"))
109           lngPaintID = Val("" & rsTmp("����ID"))
110           If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, , strErr) Then
111               If byRunMode = 3 Then
112                   FunReportPrintSet gcnHisOracle, 100, strReportCode, objFrm
113               Else
114                   If ReadSampleImage(lngSampleID, strChart, strErr, 10) = False Then
115                       Exit Function
116                   End If

117                   If strPrinter <> "" Then Call FunSetReportPrintSet(gcnHisOracle, 100, strReportCode, "printer", strPrinter)  '����ָ����ӡ��
118                   Call FunReportOpen(gcnHisOracle, 100, strReportCode, objFrm, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                                         "����ID=" & lngPaintID, "�걾ID=" & lngSampleID, _
                                         "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                                         "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                                         "ͼ��9=" & strChart(8), "PDF=" & strPDF, byRunMode)
119               End If
120           Else
121               Exit Function
122           End If
123       End If

124       If Not blnNewReport And Not blnOldReport Then
125           strErr = "δ��ѯ������"
126           Exit Function
127       End If


128       funPrintLisReport = True


129       Exit Function
funPrintLisReport_Error:
130       strErr = Err.Description
131       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funPrintLisReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
132       Err.Clear
End Function

Public Function GetReportCode(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False, Optional ByRef strErr As String) As Boolean
          '--------------------------------------------------------------------------------------------------------
          '����;
          '--------------------------------------------------------------------------------------------------------
          Dim rs As New ADODB.Recordset
          Dim strSQL As String
          
1         On Error GoTo GetReportCode_Error

2         If lngҽ��ID = 0 And lng���ͺ� = 0 Then Exit Function
          
3         strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
                             "A.NO," & _
                             "A.��¼���� " & _
                      "FROM ����ҽ������ A,�����ļ��б� C,����ҽ����¼ D,��������Ӧ�� E " & _
                      "Where E.�����ļ�id = C.ID " & _
                              "AND D.������ĿID=E.������ĿID " & _
                            "AND A.ҽ��ID=D.ID AND E.Ӧ�ó���=Decode(D.������Դ,2,2,4,4,1) " & _
                            " AND D.���id= [1] "
4         If DataMoved Then
5             strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
6             strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
7         End If

8         Set rs = ComOpenSQL(Sel_His_DB, strSQL, "�����ӡ", lngҽ��ID, lng���ͺ�)
                            
          
9         If rs.BOF = False Then
10            strCode = NVL(rs("������"))
11            strNO = NVL(rs("NO"))
12            bytMode = NVL(rs("��¼����"), 1)
13        End If
14        GetReportCode = True


15        Exit Function
GetReportCode_Error:
16        strErr = Err.Description
17        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetReportCode)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
18        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-03-13
'��    ��:  ͨ�����Ʊ����ȡ��ǰ��Ŀ�Ƿ�Ϊ����ʵ����Ŀ
'��    ��:
'           strItemID       ������ĿID
'��    ��:
'           strErr          ���������ʾ��Ϣ
'��    ��:  True=��ǰ��Ŀ��������Ŀ��False=��ǰ��Ŀ����������Ŀ
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function funIsToleranceItem(ByVal strItemID As Long, ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemCode As String       '���Ʊ���

1         On Error GoTo funIsToleranceItem_Error
          
2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") = -1 Then
3             Exit Function
4         End If
          
          'ͨ��������ĿID��ȡ������Ŀ����
5         strSQL = "select ���� from ������ĿĿ¼ where ID=[1]"
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strItemID)
7         If rsTmp.EOF Then
      '        strErr = "û�в�ѯ����Ӧ��������Ŀ"
8             Exit Function
9         Else
10            strItemCode = Trim(rsTmp("����") & "")
11        End If

          '�ж��Ƿ�����������
12        strSQL = "select id,�Ƿ�������Ŀ from ���������Ŀ where ���Ʊ���=[1]"
13        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "������Ŀ", strItemCode)
14        If rsTmp.RecordCount > 1 Then
15            strErr = "��ǰ��Ŀ�����˶��������Ŀ������ϵ����������Ա�����Ų�"
16            Exit Function
      '    ElseIf rsTmp.RecordCount = 0 Then
      '        strErr = "��ǰ��Ŀδ���°�LIS��Ŀ���ж���"
17        ElseIf rsTmp.RecordCount = 1 Then
18            If Val(rsTmp("�Ƿ�������Ŀ") & "") = 1 Then
19                funIsToleranceItem = True
20            End If
21        End If


22        Exit Function
funIsToleranceItem_Error:
23        strErr = "ִ��(funIsToleranceItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl
24        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funIsToleranceItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
25        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-15
'��    ��:  ͨ������ҽ��ID����������Щҽ��������ͬһ���걾
'��    ��:
'           strAdvice       ҽ��ID�����ҽ��IDʹ��Ӣ�Ķ��ŷָ�
'��    ��:
'��    ��:  ҽ��ID������ͬ�걾֮���ҽ��IDʹ��Ӣ�ķֺŷָ��ͬ�걾��ҽ��IDʹ�ö��ŷָ�
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function funGetSampleAdvice(ByVal strAdvice As String) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strOld As String
          Dim strReturn As String
          Dim strArr() As String
          Dim i As Integer

1         On Error GoTo funGetSampleAdvice_Error

2         strOld = "," & strAdvice & ","
3         strArr = TruncatedExtraLongStr(strAdvice, ",")
4         For i = 0 To UBound(strArr)
5             strOld = "," & strArr(i) & ","
              '���°���ȥ����
6             strSQL = "Select /*+cardinality(b,10)*/ f_List2str(Cast(Collect(a.����id || '') As t_Strlist)) ����id" & vbCrLf & _
                     " From ����������� A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                     " Where A.����id = B.Column_Value and a.�걾ID is not null" & vbCrLf & _
                     " Group By a.�걾ID"
7             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", strArr(i))
              '�޳��°��ҽ��ID֮��ʣ�µĵ��ϰ�LIS��ȥ��ѯ
8             Do While Not rsTmp.EOF
9                 If Not IsNull(rsTmp("����id")) Then
10                    strOld = Replace(strOld, rsTmp("����id") & ",", "")
11                    If InStr(rsTmp("����id") & "", ",") > 0 Then
12                        strReturn = strReturn & ";" & rsTmp("����id")
13                    End If
14                End If
15                rsTmp.MoveNext
16            Loop
17            If strOld <> "" Then
18                If Left(strOld, 1) = "," Then strOld = Mid(strOld, 2)
19                If Right(strOld, 1) = "," Then strOld = Mid(strOld, 1, Len(strOld) - 1)
20            End If
21            If strOld <> "" Then
                  '���ϰ���ȥ����
22                strSQL = "Select /*+cardinality(b,10)*/ f_List2str(Cast(Collect(a.ҽ��id || '') As t_Strlist)) ����id" & vbCrLf & _
                         " From ������Ŀ�ֲ� A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                         " Where A.ҽ��id = B.Column_Value" & vbCrLf & _
                         " Group By a.�걾ID"
23                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�����������", strOld)
24                Do While Not rsTmp.EOF
25                    If Not IsNull(rsTmp("����id")) Then
26                        If InStr(rsTmp("����id") & "", ",") > 0 Then
27                            strReturn = strReturn & ";" & rsTmp("����id")
28                        End If
29                    End If
30                    rsTmp.MoveNext
31                Loop
32            End If
33        Next
34        If strReturn <> "" Then
35            If Left(strReturn, 1) = ";" Then funGetSampleAdvice = Mid(strReturn, 2)
36        End If

37        Exit Function
funGetSampleAdvice_Error:
38        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetSampleAdvice)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
39        Err.Clear
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, "ZLLIS"
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long, Optional ByVal int���� As Integer) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnHisOracle, gSysInfo.SysNo, gSysInfo.ModlNo, int����)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-19
'��    ��:  ��ʾ���Ʋο�
'��    ��:
'           objFrm          ��������
'           lngSampleID     �걾ID
'           lngVer          �汾
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Sub funShowClincHelp(objFrm As Object, ByVal lngSampleID As Long, ByVal lngVer As Long)
          Dim objAdvice As Object
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemCode As String
          Dim strItemIDs As String
          Dim lngPaitID As Long
          Dim lngPage As Long
          Dim intPaitType As Integer
          Dim strGHNo As String
          Dim lngGHID As Long
          Dim blnContinue As Boolean


1         On Error GoTo funShowClincHelp_Error

2         If lngSampleID <> 0 Then
              '��ȡ������ĿID
3             If lngVer = 25 Then
4                 strSQL = "Select f_List2str(Cast(Collect(b.���Ʊ��� || '') As t_Strlist)) ����" & vbCrLf & _
                           "   From ����������� A, ���������Ŀ B" & vbCrLf & _
                           "   Where A.���ID = b.id And a.�걾id = [1]"
5                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����������", lngSampleID)
6                 If Not rsTmp.EOF Then
7                     strItemCode = rsTmp("����") & ""
8                 End If

9                 If strItemCode <> "" Then
                      'ͨ�����Ʊ����ѯ������ĿID
10                    strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                               "f_List2str(Cast(Collect(a.ID || '') As t_Strlist)) ID" & vbCrLf & _
                               " From ������ĿĿ¼ A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                               " Where A.���� = B.Column_Value"
11                    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strItemCode)
12                    If Not rsTmp.EOF Then strItemIDs = rsTmp("ID") & ""
13                End If

                  '��ȡ������Ϣ
14                strSQL = "select ����ID,������Դ,��ҳID,�Һŵ� from ���鱨���¼ where ID=[1]"
15                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鱨���¼", lngSampleID)
16                If Not rsTmp.EOF Then
17                    lngPaitID = Val(rsTmp("����ID") & "")
18                    lngPage = Val(rsTmp("��ҳID") & "")
19                    intPaitType = Val(rsTmp("������Դ") & "")
20                    strGHNo = rsTmp("�Һŵ�") & ""
21                End If
22            ElseIf lngVer = 10 Then
23                strSQL = " select f_List2str(Cast(Collect(b.������ĿID || '') As t_Strlist)) ������ĿID from ����걾��¼ A, ����ҽ����¼ B where a.ҽ��id=b.���id and a.id=[1]"
24                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿID", lngSampleID)
25                If Not rsTmp.EOF Then
26                    strItemIDs = rsTmp("������ĿID") & ""
27                End If

                  '��ȡ������Ϣ
28                strSQL = "select ����ID,������Դ,��ҳID,�Һŵ� from ����걾��¼ where ID=[1]"
29                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鱨���¼", lngSampleID)
30                If Not rsTmp.EOF Then
31                    lngPaitID = Val(rsTmp("����ID") & "")
32                    lngPage = Val(rsTmp("��ҳID") & "")
33                    intPaitType = Val(rsTmp("������Դ") & "")
34                    strGHNo = rsTmp("�Һŵ�") & ""
35                End If

36            End If

              '��ѯ�Һ�ID
37            If strGHNo <> "" And lngPaitID <> 0 Then
38                strSQL = "SELECT ID FROM ���˹Һż�¼ where no=[1] AND ����ID=[2]"
39                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���˹Һż�¼", strGHNo, lngPaitID)
40                If Not rsTmp.EOF Then
41                    lngGHID = Val(rsTmp("ID") & "")
42                End If
43            End If
44        End If

          '�ȵ���plugin�еĽӿڣ��ӿڵ���ʧ���ٵ���zlPublicAdvice�еĽӿ�
45        If VerCompare(gSysInfo.VersionHIS, "10.35.130") <> -1 Then
46            If CreatePlugInOK(2500, 2) Then
47                On Error Resume Next
48                blnContinue = gobjPlugIn.ShowClinicHelp(objFrm.hWnd, 1, intPaitType, lngPaitID, IIf(intPaitType = 2, lngPage, lngGHID), strItemIDs)
49                Call zlPlugInErrH(Err, "ExecuteFunc")
50                Err.Clear: On Error GoTo 0
51            End If
52        End If

          '���ýӿ�
53        If Not blnContinue Then
54            If objAdvice Is Nothing Then
55                Set objAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
56                If Not objAdvice Is Nothing Then
57                    On Error Resume Next
58                    Call objAdvice.ShowClincHelp(1, objFrm, 0, False, strItemIDs)
59                    If Err.Number = 438 Then
60                        MsgBox "HIS�汾����", vbInformation, gSysInfo.AppName
61                        Exit Sub
62                    End If
63                End If
64            End If
65        End If



66        Exit Sub
funShowClincHelp_Error:
67        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funShowClincHelp)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
68        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-06-27
'��    ��:  ������Ҳ����ť
'��    ��:
'           objCbr          CommandBar����
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function CreatePlugInButton(objToolBar As CommandBar) As Boolean
          Dim cbrMenuBar As CommandBarPopup
          Dim cbrControl As CommandBarControl
          Dim cbrToolBar As CommandBar
          Dim strTmp As String
          Dim arrTmp As Variant
          Dim i As Integer

          '-----------------------------------------��Ӳ��-------------------------------------------------
          '�����չ����
1         On Error GoTo CreatePlugInButton_Error

2         Call CreatePlugInOK(2500, 1)
3         If Not gobjPlugIn Is Nothing Then
4             With objToolBar.Controls
5                 Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Tool_PlugIn, "��չ����(&G)")
6                 cbrControl.BeginGroup = True
7                 cbrControl.Style = xtpButtonIconAndCaption
8                 With cbrControl.CommandBar.Controls
9                     If Not gobjPlugIn Is Nothing Then
10                        On Error Resume Next
11                        strTmp = gobjPlugIn.GetFuncNames(2500, 2500, 2)
12                        Call zlPlugInErrH(Err, "GetFuncNames")
13                        Err.Clear: On Error GoTo 0
14                    End If
15                    If strTmp <> "" Then
16                        strTmp = Replace(strTmp, "Auto:", "")
17                        arrTmp = Split(strTmp, ",")
18                        For i = 0 To UBound(arrTmp)
19                            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
20                            If i <= 9 Then cbrControl.Caption = cbrControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
21                            cbrControl.IconId = conMenu_Tool_PlugIn_Item
22                            cbrControl.Parameter = arrTmp(i)
23                        Next
24                    End If
25                End With
26            End With
27        End If
          '-----------------------------------------END-------------------------------------------------


28        Exit Function
CreatePlugInButton_Error:
29        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(CreatePlugInButton)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
30        Err.Clear
End Function

Public Sub ExePlugIn(ByVal strName As String, ByVal lngSampleID As Long)
'���ܣ�ִ����ҹ���
    Dim lngID As String
    Dim lngPaitID As Long
    Dim lngMainID As Long
    If CreatePlugInOK(2500, 1) Then
        Call gobjPlugIn.ExecuteFunc(2500, 2500, strName, lngPaitID, lngMainID, lngID, lngSampleID, 1)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-06-27
'��    ��:  ��ȡ�걾���䱨��
'��    ��:
'           lngSampleID     �걾ID
'           objVSF          չʾ���ݵ�VSF
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function GetSupplementReport(ByVal lngSampleID As Long, objVSF As VSFlexGrid) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim intItem As Integer
          Dim strItem As String
          Dim lngRow As Long
          
1         On Error GoTo GetSupplementReport_Error

2         intItem = ComGetPara(Sel_Lis_DB, "������Ŀ��ʾ", gSysInfo.SysNo, gSysInfo.ModlNo, "1")
          
          
3         Select Case intItem
              Case 1
4                 strItem = "c.������  ������Ŀ"

5             Case 2
6                 strItem = "c.Ӣ����  ������Ŀ"

7             Case 3
8                 strItem = "c.������ || '(' || c.Ӣ���� || ')'  ������Ŀ"
9             End Select

10            strSQL = "Select b.id, b.���䱨��ID, b.��ĿID," & strItem & ", b.����ID, b.������, b.�����־, b.����ο�, b.�ο���ֵ, b.�ο���ֵ, b.��λ" & vbCrLf & _
                      " From ���鲹�䱨���¼ A, ���鲹�䱨����ϸ B, ����ָ�� C" & vbCrLf & _
                      " Where a.ID = b.���䱨��id And b.��Ŀid = c.ID And a.�걾id = [1] And b.���������� = 2" & vbCrLf & _
                      " Order By c.�������"
11            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���䱨����ϸ", lngSampleID)
12            If SetDataToVSF(objVSF, rsTmp) = False Then Exit Function
13            With objVSF
14                .SelectionMode = flexSelectionFree
15                .ColHidden(.ColIndex("ID")) = True
16                .ColHidden(.ColIndex("���䱨��ID")) = True
17                .ColHidden(.ColIndex("��ĿID")) = True
18                .ColHidden(.ColIndex("����ID")) = True
19                .ColHidden(.ColIndex("�����־")) = True
20                .ColHidden(.ColIndex("�ο���ֵ")) = True
21                .ColHidden(.ColIndex("�ο���ֵ")) = True
                  
22                For lngRow = 1 To .Rows - 1
23                    .Cell(flexcpBackColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = GetValColour(Val(.TextMatrix(lngRow, .ColIndex("�����־"))))
24                Next
25            End With
              

26        Exit Function
GetSupplementReport_Error:
27        Call WriteErrLog("ZL9LabWork", "mdlWorkBaseReprot", "ִ��(GetSupplementReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
28        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-06-27
'��    ��:  ������б��е�����ָ���Ϊɾ����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Sub EditSampleValueList(objVSFSampleValue As VSFlexGrid, objVSFSupplement As VSFlexGrid)
    Dim i As Integer
    Dim J As Integer
    
    With objVSFSampleValue
        For i = 1 To .Rows - 1
            With objVSFSupplement
                For J = 1 To .Rows - 1
                    If Val(objVSFSampleValue.TextMatrix(i, objVSFSampleValue.ColIndex("ID"))) = Val(.TextMatrix(J, .ColIndex("��ĿID"))) Then
                        objVSFSampleValue.Cell(flexcpFontStrikethru, i, 0, i, objVSFSampleValue.Cols - 1) = True
                    End If
                Next
            End With
        Next
    End With
End Sub

Public Function GetValColour(intValType As Integer) As Double
    '����               �����Ӧ�Ľ������1-������2-ƫ�͡�3-ƫ�ߡ�4-����(�쳣)��5-��ʾ���ޡ�6-��ʾ���ޡ�7-�������ޡ�8-��������
    '����               ��Ӧ����ɫ
    Select Case intValType
        Case 1, 0
            GetValColour = gSampleShowColour.����
        Case 2
            GetValColour = gSampleShowColour.ƫ��
        Case 3
            GetValColour = gSampleShowColour.ƫ��
        Case 4
            GetValColour = gSampleShowColour.�쳣
        Case 5
            GetValColour = gSampleShowColour.��ʾƫ��
        Case 6
            GetValColour = gSampleShowColour.��ʾƫ��
        Case 7
            GetValColour = gSampleShowColour.����ƫ��
        Case 8
            GetValColour = gSampleShowColour.����ƫ��
    End Select
End Function



'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-09-26
'��    ��:  ͨ������ID��ȡ�����Ŀ��ϸ
'��    ��:
'           strInfo         intType=0,�����Ŀ��Ӧ��������Ŀ��ID�����ʹ�á�,���ָ�;intType=1,���Ʊ��룬���ʹ�ö��ŷָ�
'��    ��:
'��    ��:  ��������Ŀ��ָ���¼��
'����Ӱ��:
'����ע��:  intType=0ʱ�����صļ�¼����4���ֶ�:ָ��ID��ָ�����ƣ����ܷ���ID������ʱ��
'           intType=1ʱ�����صļ�¼����8���ֶ�:������ƣ���ϱ��룬���Ʊ��룬�걾���ͣ�ָ��ID��ָ�����ƣ����ܷ���ID������ʱ��
'---------------------------------------------------------------------------------------
Public Function funGetGroupItemInfo(ByVal strInfo As String, Optional ByVal intType As Integer) As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strCode As String

          'ͨ��������ĿID��ѯ������Ŀ����
1         On Error GoTo funGetGroupItemInfo_Error

2         If intType = 0 Then
              'ͨ��������ĿID��ȡ�����ϸ
3             strSQL = "Select /*+cardinality(d,10)*/ f_List2str(Cast(Collect(a.����) As t_Strlist)) ����" & vbCrLf & _
                       "   From ������ĿĿ¼ A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) B" & vbCrLf & _
                       "   Where a.id = b.Column_Value"
4             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strInfo)
5             If Not rsTmp.EOF Then
6                 strCode = rsTmp("����") & ""
7             End If

              'ͨ��������Ŀ�����ѯ���������ϸ
8             If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
9                 strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "   c.ID ָ��ID, c.������ ����ָ��, d.id skey, d.����ʱ�� sname" & vbCrLf & _
                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C, ��������ʱ�䷽�� D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where a.id = b.���id And b.��Ŀid = c.id And b.��Ŀid = d.��Ŀid(+) And a.���Ʊ��� = e.Column_Value"
10            Else
11                strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "   c.ID ָ��ID, c.������ ����ָ��, '' skey, '' sname" & vbCrLf & _
                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where a.id = b.���id And b.��Ŀid = c.id And a.���Ʊ��� = e.Column_Value"
12            End If
13            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����Ŀ��ϸ", strCode)
14        ElseIf intType = 1 Then
              'ͨ��������Ŀ�����ѯ���������ϸ
15            If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
16                strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "    a.����, a.����, a.���Ʊ���, a.����걾 �걾����, c.ID ָ��ID, c.������ ����ָ��, d.id ���ܷ���ID, d.����ʱ��" & vbCrLf & _
                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C, ��������ʱ�䷽�� D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where a.id = b.���id And b.��Ŀid = c.id And b.��Ŀid = d.��Ŀid(+) And a.���Ʊ��� = e.Column_Value" & vbCrLf & _
                           "   Order By a.id, c.�������"
17            Else
18                strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "    a.����, a.����, a.���Ʊ���, a.����걾 �걾����, c.ID ָ��ID, c.������ ����ָ��, '' ���ܷ���ID, '' ����ʱ��" & vbCrLf & _
                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where A.ID = B.���id And B.��Ŀid = C.ID And A.���Ʊ��� = e.Column_Value" & vbCrLf & _
                           "   Order By a.id, c.�������"
19            End If
20            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����Ŀ��ϸ", strInfo)
21        End If
22        Set funGetGroupItemInfo = rsTmp


23        Exit Function
funGetGroupItemInfo_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(funGetGroupItemInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-10-29
'��    ��:  ͨ��ҽ��ID��ȡXML��ʽ�Ĳ��˱��棨�°�LIS���ϰ�LIS��
'��    ��:
'           strAdviceID     ҽ��ID�������ʹ�ö��ŷָ�
'��    ��:
'           strErr          ������Ϣ������ʾ��Ϣ
'��    ��:  ��������ҽ��ID��Ӧ�����м���ָ�꼰�������¼�����ݣ�������,������,��λ

'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function funGetPatientReport(ByVal strAdviceID As String, Optional ByRef strErr As String) As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsData As ADODB.Recordset
          Dim strXML As String

          '��ѯ�°�LIS����
1         On Error GoTo funGetPatientReport_Error

2         strSQL = "Select Distinct /*+cardinality(e,10)*/  '[' || f.���� || ']' || c.������ ������, a.������, c.��λ" & vbCrLf & _
                 "   From ���鱨����ϸ A, ����������� B, ����ָ�� C, ���鱨���¼ D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E, ���������Ŀ F" & vbCrLf & _
                 "   Where b.ҽ��id = e.Column_Value And a.�걾id = b.�걾id And a.��Ŀid = c.id And a.�걾id = d.ID And a.���id = f.id And" & vbCrLf & _
                 "         d.����� Is Not Null"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�°���", strAdviceID)

4         Set rsData = gobjHisDatabase.CopyNewRec(rsTmp)

          '��ѯ�ϰ�
5         strSQL = "Select Distinct /*+cardinality(e,10)*/  '[' || f.���� || ']' || c.������ ������, b.������, c.��λ" & vbCrLf & _
                 "   From ������Ŀ�ֲ� A, ������ͨ��� B, ����������Ŀ C, ����걾��¼ D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E, ������ĿĿ¼ F" & vbCrLf & _
                 "   Where a.�걾id = b.����걾id And a.��Ŀid = b.������Ŀid And b.������Ŀid = c.id And a.�걾ID = d.id And b.������ĿID = f.id And" & vbCrLf & _
                 "         a.ҽ��Id = e.Column_Value And d.����� Is Not Null"
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�ϰ���", strAdviceID)

          '���ϰ���°汨��ϲ���һ��
7         Do While Not rsTmp.EOF
8             rsData.AddNew
9             rsData("������") = rsTmp("������") & ""
10            rsData("������") = rsTmp("������") & ""
11            rsData("��λ") = rsTmp("��λ") & ""

12            rsTmp.MoveNext
13        Loop
14        If rsData.RecordCount > 0 Then rsData.MoveFirst

15        Set funGetPatientReport = rsData

16        Exit Function
funGetPatientReport_Error:
17        strErr = WriteErrLog("zl9LisInsideComm", "mdlLisHisComm", "ִ��(funGetPatientReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
18        Err.Clear
End Function
