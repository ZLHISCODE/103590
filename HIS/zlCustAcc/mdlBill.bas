Attribute VB_Name = "mdlBill"
Option Explicit

'��ģ����ר��Ϊ�շѼ��ʶ�������
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public gbln�����л� As Boolean
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    pסԺ���� = 1133
    p���˽��� = 1137
    p���ò�ѯ = 1139
    pһ���嵥 = 1141
    p���ʲ��� = 1150
End Enum

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    
    supportʵʱ��� = 60             '�Ƿ����÷���ʵʱ���
End Enum

'ϵͳ������ʱ����
'Public glngOK As Long   '����ɹ��ͷ���1��ȡ���ͷ���0����������ڼ��ʵ�ģ�屻ɾ��������-1
Public gblnOK As Boolean
Public gbytWarn As Byte '���ʱ�������ֵ
Public gstrModiNO As String '�޸ĺ�������µ��ݺ�

'============ҽ������=====================
Public gclsInsure As New clsInsure
'============����ϵͳ����=====================
Public grsPar As ADODB.Recordset '��¼ϵͳ����

'ˢ������
Public gbytCardNOLen As Byte '���￨�ų���
Public gblnShowCard As Boolean '�Ƿ���￨����ʾΪ��������
Public gstrCardPass As String 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'

'�����������
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gbytCode As Byte '�������ɷ�ʽ��0-ƴ��,1-���,2-����
Public gblnMyStyle As Boolean 'ʹ�ø��Ի����
Public gstrIme As String '�Զ��Ŀ������뷨

Public gstrMatchMode As String '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
Public gbytҽ�������� As Byte '0-�����м�顢1-��鲢����δ������Ŀ��2-��鲢��ֹδ������Ŀ
Public gcurMaxMoney As Currency '���ʷ���������ѽ��

'��������
Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��

'��������
Public gblnDailyTime As Boolean '��ʱ��ʾ����ʱ�䳬��������ʱ��
Public gbln�����������۷��� As Boolean '���ʱ����������۷���

Public gintOutDay As Integer '���ʿ�ѡ���Ժ��������

'�������
Public gstr�շ���� As String '��������շ����
Public gblnTime As Boolean '����Ƿ������������
Public gbln��ʿ As Boolean '�������Ƿ���ʾ��ʿ
Public gbln������ As Boolean '�����Ƿ�������뿪����

'���۲��˼���
Public gbln�������� As Boolean
Public gblnסԺ���� As Boolean
'���˺� ����:????    ����:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000

'���С��λ��
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

'ϵͳ����:34717
Public Type TY_Reg_Para  '�Һ���ز���
    bytNODaysGeneral As Byte    '��ͨ�Һ���Ч����
    bytNoDayseMergency As Byte '����Һ���Ч����
End Type
Public Type TY_SysPara
    Sy_Reg  As TY_Reg_Para
    byt������˷�ʽ As Byte '49501:������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
    blnδ��ƽ�ֹ���� As Boolean '51612
End Type
Public gSysPara As TY_SysPara       'ϵͳ�������;�Ժ������չ(���˺�)

Private mlng���ű���ƽ������ As Long

Public Function BillingWarn(strPrivs As String, str���� As String, lng����ID As Long, str���ò��� As String, _
    rsWarn As ADODB.Recordset, cur��� As Currency, cur���ն� As Currency, _
    cur���ݽ�� As Currency, cur���� As Currency, str��� As String, _
    ByVal str����� As String, ByRef str�ѱ���� As String, Optional bln�ಡ�� As Boolean, _
    Optional curItemMoney As Currency = 0, _
    Optional blnNotCheck��� As Boolean = False) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:
'     str����=��������,������ʾ
'     lng����ID=���˲���ID,����ѡ�����õĲ����������ã�0��ʾû��ȷ�������������в����ı�����������
'     str���ò���=���ݲ�����ݷ��صļ��ʱ������÷���
'     rsWarn=��ǰ�������ʱ������ü�¼
'     cur���=�������,�����ۼƱ���
'     cur���ն�=���˵��շ����ķ��ö�,����ÿ�ձ���
'     cur���ݽ��=���˵���������ķ���
'     cur����=���˵������ö�,�����ۼƱ���
'     str���=��ǰҪ�������,���ڷ��౨��
'     str�����=�������,������ʾ
'     curItemMoney-���ʽ��(�������<>0 ,����Ҫ�жϵ������,����������,�������û�����,������ݱ�����ʽ����):���˺�:24491
'     blnNotCheck���:���������м��(��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
'����:0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
'     str�������="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
    Dim i As Integer, byt��־ As Byte
    Dim bln�ѱ��� As Boolean
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    '�����������
    If rsWarn.State = 0 Then Exit Function '20030709
    rsWarn.Filter = "���ò���='" & str���ò��� & "' And ����ID=" & lng����ID
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str���) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str����� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־1 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str���) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str����� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־2 <> "-" And blnNotCheck��� Then Exit Function
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str���) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str����� = "" '�������ʱ,������ʾ��������
        '���˺� ����:26952 ����:2009-12-25 16:42:54
        '   ��Ҫ������Ը�ѡ���˺󣬻�δ������ص�����ʱ���״μ��.�����ֻ��������Ƶ����Ϊ������������������Ƶģ�����������¾Ͳ����,ֻ�����������ݺ�ż��!)
        If rsWarn!������־3 <> "-" And blnNotCheck��� Then Exit Function
    End If
    
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    If bln�ಡ�� Then
        'ʾ����",��:-,��:DEF,��:567,��567"
        '������־2ʾ����",��:-��,��:DEF��,��:567��,��567��"
        bln�ѱ��� = str�ѱ���� & "," Like "*," & str���� & ":-*,*" _
            Or str�ѱ���� & "," Like "*," & str���� & ":*" & str��� & "*,*"
    Else
        'ʾ����"-" �� ",ABC,567,DEF"
        '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
        bln�ѱ��� = InStr(str�ѱ����, str���) > 0 Or str�ѱ���� Like "-*"
    End If
    
    If bln�ѱ��� Then
        If byt��־ = 2 Then
            If bln�ಡ�� Then
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str���� & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str���� & ":*" & str��� & "*,*" Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        Exit For
                    End If
                Next
            Else
                If str�ѱ���� Like "-*" Then
                    byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
                Else
                    arrTmp = Split(str�ѱ����, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str���) > 0 Then
                            byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str����� <> "" Then str����� = """" & str����� & """����"
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ", �Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                    End If
                
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1
                        End If
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & " ����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If cur��� + cur���� - cur���ݽ�� < 0 Then
                        '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                         If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > 0 Then
                             'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                             '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                            If MsgBox("ע��" & vbCrLf & _
                                       "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�,�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                                             
                        byt��ʽ = 2
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                            BillingWarn = 3
                        Else
                            MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    ElseIf cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                        '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                         If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > Val(Nvl(rsWarn!����ֵ)) Then
                             'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                             '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                            If MsgBox("ע��" & vbCrLf & _
                                       "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ", �Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                    
                    
                        byt��ʽ = 1
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1
                            End If
                        Else
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If cur��� + cur���� - cur���ݽ�� < 0 Then
                            '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                             If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > 0 Then
                                'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                                '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                               If MsgBox("ע��" & vbCrLf & _
                                          "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�,�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                   BillingWarn = 2
                               Else
                                   BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                               End If
                               Exit Function
                            End If
                         
                            byt��ʽ = 2
                            If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                                MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                                BillingWarn = 3
                            Else
                                MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur��� + cur���� - (cur���ݽ�� - curItemMoney) > Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�,�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                    
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur���ն� + cur���ݽ�� - curItemMoney < Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        If MsgBox(str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1
                        End If
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    '24491 ���˺�:��ʾԤ����۳����ʽ���Ѿ��ľ�  ,��ԭ���Ĺ���������,����ֻ��ʾ
                     If curItemMoney <> 0 And cur���ն� + cur���ݽ�� - curItemMoney < Val(Nvl(rsWarn!����ֵ)) Then
                         'ֻ��ʾ: gbytBilling As Byte '0-����,1-����,2-���
                         '��Ҫ�Ǽ���������:����¼�뵱����Ŀʱ��Ҫ�������Ρ�������Ŀ�ȣ��Ӷ�����ʵ�ս��ļ��٣����в��������㱨������
                        If MsgBox("ע��" & vbCrLf & _
                                   "    ���ˡ�" & str���� & "�� ���շ���:" & Format(cur���ն� + cur���ݽ��, gstrDec) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",�Ƿ����¼�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        MsgBox str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetDoctor(lng����ID As Long, ByVal bln��ʿ As Boolean, ByRef rsTmp As ADODB.Recordset, ByVal int������Դ As Integer)
'���ܣ���ȡָ�����ҵ�ҽ��
'������lng����ID=ָ������ID,bln��ʿ=�Ƿ�Ҳ��ȡ��ʿ(�շ�\����)
    'Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '�����Ź�������Ϊ���ٴ���ҽ����ʿ,��Ϊ���ܸ���������һ�������Ƿ�ĩ������.
    If rsTmp Is Nothing Then
        strSQL = _
            "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
            " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��,B.ȱʡ" & _
            " From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D" & _
            " Where A.ID = B.��ԱID And A.ID=C.��ԱID And B.����ID=D.����ID And C.��Ա���� IN('ҽ��','��ʿ') " & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " And D.������� IN(" & int������Դ & ",3) And D.�������� IN('�ٴ�','����') And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
            " Order by ����,ȱʡ Desc"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlBill")
    End If
   
    If lng����ID = 0 Then
        rsTmp.Filter = IIf(bln��ʿ, "", "��Ա����='ҽ��'")
    Else
        rsTmp.Filter = "����ID=" & lng����ID & IIf(bln��ʿ, "", " And ��Ա����='ҽ��'")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Public Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str������ As String, ByVal str�������� As String, _
    Optional ByVal lngRow As Long = -1) As ADODB.Recordset
'���ܣ����ݵ��ݶ������ݴ���һ����ϸ��¼����Ϣ(���ۼ۵�λ)
'�ֶΣ�����ID����ҳID���շ�����շ�ϸĿID�����������ۣ�ʵ�ս������ˣ���������
'������intPage=ָ���ĵ���,lngRow=ָ�����У���ָ��ʱ�������е��ݵ�������,ע���к��Ǵ�0��ʼ�ģ����Ҷ����м���Rǰ׺
    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl���� As Double, curʵ�� As Currency
    Dim rsTmp As New ADODB.Recordset
    '79420,���ϴ�,2014/11/10:������¼���ֶδ�С
    rsTmp.Fields.Append "����ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "�շ����", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "����", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "ʵ�ս��", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    If lngRow = -1 Then
        intB = 0
        intE = objBill.Details.Count - 1
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl���� = 0: curʵ�� = 0
        With objBill.Details("R" & i)
            If .�շ�ϸĿID <> 0 Then    '�����objBill�е���ϸ���ܻ�û������Ϣ����Ԥ��Ƶ�����
                If lngRow = -1 Then
                    rsTmp.Filter = "�շ�ϸĿID=" & .�շ�ϸĿID
                    blnNew = rsTmp.RecordCount = 0
                Else
                    blnNew = True
                End If
                                
                If blnNew Then
                    rsTmp.AddNew
                    
                    rsTmp!����ID = objBill.����ID
                    rsTmp!��ҳID = objBill.��ҳID
                    
                    rsTmp!�շ���� = .�շ����
                    rsTmp!�շ�ϸĿID = .�շ�ϸĿID
                    
                    
                    For j = 1 To .InComes.Count
                        dbl���� = dbl���� + .InComes(j).��׼����
                        curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                    Next
                    rsTmp!���� = .����
                    rsTmp!���� = Format(dbl����, gstrFeePrecisionFmt)
                    
                    rsTmp!ʵ�ս�� = Format(curʵ��, gstrDec)
                    
                    rsTmp!������ = str������
                    rsTmp!�������� = str��������
                Else
                    For j = 1 To .InComes.Count
                        dbl���� = dbl���� + .InComes(j).��׼����
                        curʵ�� = curʵ�� + .InComes(j).ʵ�ս��
                    Next
                    rsTmp!���� = rsTmp!���� + .����
                    rsTmp!���� = Format((rsTmp!���� + Format(dbl����, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                    
                    rsTmp!ʵ�ս�� = rsTmp!ʵ�ս�� + Format(curʵ��, gstrDec)
                End If
                
                rsTmp.Update
            End If
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-���˱���"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call zlControl.CboSetIndex(cboBaby.hwnd, 0)
    
    If lngPatient <> 0 Then
        Set rsTmp = GetPatientBaby(lngPatient, lngPatientPage)
        With rsTmp
            For i = 1 To .RecordCount
                If Not IsNull(!Ӥ������) Then
                    cboBaby.AddItem !��� & "-" & !Ӥ������
                Else
                    cboBaby.AddItem !��� & "-��" & !��� & "��Ӥ��"
                End If
                cboBaby.ItemData(cboBaby.NewIndex) = !���
                .MoveNext
            Next
        End With
    End If
End Sub

Public Function GetPatientBaby(ByVal lngPatient As Long, lngPatientPage As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select ���, Ӥ������ From ������������¼ Where ����id = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������¼", lngPatient, lngPatientPage)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim strValue As String
    
    On Error Resume Next
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '������ʾ��ʽ
    gblnShowCard = zlDatabase.GetPara(12, glngSys) = "0"
    
   '���￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytCardNOLen = Val(Split(strValue, "|")(4))
    '����:35242
    gbln�����л� = IIf(Val(zlDatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1)) = 1, 1, 0) = 1
        
    '�Һ���Ч����
    '���˺�:34717
    '��λ:ǰһλ���ܹҺ�;��һλ����Һ�
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    gSysPara.Sy_Reg.bytNODaysGeneral = Val(Left(strValue, 1))
    gSysPara.Sy_Reg.bytNoDayseMergency = Val(Right(strValue, 1))
    'If gSysPara.Sy_Reg.bytNODaysGeneral = 0 Then gSysPara.Sy_Reg.bytNODaysGeneral = 1
    ' If gSysPara.Sy_Reg.bytNoDayseMergency = 0 Then gSysPara.Sy_Reg.bytNoDayseMergency = 1
    '49501
    gSysPara.byt������˷�ʽ = Val(zlDatabase.GetPara(185, glngSys, , "0"))
    gSysPara.blnδ��ƽ�ֹ���� = Val(zlDatabase.GetPara(215, glngSys, , "0")) = 1 '51612
    '�ձ�ͳ��ʱ������
    gblnDailyTime = zlDatabase.GetPara(22, glngSys) = "1"
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    
    '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    gstrMatchMode = zlDatabase.GetPara(44, glngSys)
                
    'ˢ��Ҫ����������
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
    gbln������ = zlDatabase.GetPara(52, glngSys) = "1"
    gbytҽ�������� = Val(zlDatabase.GetPara(59, glngSys))
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    gbln�����������۷��� = zlDatabase.GetPara(98, glngSys) = "1"
    '���˺� ����:????    ����:2010-12-06 23:38:53
    '���õ��۱���λ��
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    
    InitSysPar = True
End Function

Public Sub InitLocPar(bytUseType As Byte)
'���ܣ���ʼ�����ñ�������
'������bytUseType=0-סԺ����,1-��ɢ����,2-ҽ������,3-����,-1-����
    Dim strValue As String
    
    If bytUseType = -1 Then Exit Sub
    
    '����ƥ������
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = 0, "%", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"

    
    '���ʿ�ѡ���Ժ��������
    gintOutDay = Val(zlDatabase.GetPara("��Ժ��������", glngSys, glngModul))
    
    gblnTime = zlDatabase.GetPara("�������", glngSys, glngModul) = "1"
    If bytUseType = 3 Then
        gstr�շ���� = zlDatabase.GetPara("�շ����", glngSys, 1121)
    Else
        gstr�շ���� = zlDatabase.GetPara("�շ����", glngSys, 1150)
        gbln��ʿ = zlDatabase.GetPara("��ʾ��ʿ", glngSys, glngModul) = "1"
    End If
    
    '���۲��˼���
    If bytUseType <> 3 Then
        gbln�������� = (zlDatabase.GetPara("�������۲��˼���", glngSys, 1150, "0") = "1")
        gblnסԺ���� = (zlDatabase.GetPara("סԺ���۲��˼���", glngSys, 1150, "0") = "1")
    End If
End Sub

Public Sub GetBillDeptID(strNO As String, lng����ID As Long, lngִ��ID As Long, ByVal int��Դ As Integer)
'���ܣ���ȡһ�ż��ʵ��ݵĿ������Һ�ִ�п���ID
'int��Դ:1-����;2-סԺ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��������ID,ִ�в���ID From " & IIf(int��Դ = 1, "������ü�¼", "סԺ���ü�¼") & " Where ��¼����=2 And ��¼״̬ IN(1,3) And NO=[1] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlBill", strNO)
    If Not rsTmp.EOF Then
        lng����ID = IIf(IsNull(rsTmp!��������ID), 0, rsTmp!��������ID)
        lngִ��ID = IIf(IsNull(rsTmp!ִ�в���ID), 0, rsTmp!ִ�в���ID)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function GetBillTotal(objBill As ExpenseBill) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillDetail As BillDetail
    Dim objBillIncome As BillInCome
    
    For Each objBillDetail In objBill.Details
        For Each objBillIncome In objBillDetail.InComes
            GetBillTotal = GetBillTotal + objBillIncome.ʵ�ս��
        Next
    Next
End Function
Public Function GetBillRowTotal(objBillInComes As BillInComes) As Currency
'���ܣ���ȡ���ݷ�Ŀ�ϼƽ��
    Dim objBillIncome As New BillInCome
    For Each objBillIncome In objBillInComes
        GetBillRowTotal = GetBillRowTotal + objBillIncome.ʵ�ս��
    Next
End Function

Public Function CheckScope(curL As Currency, curR As Currency, curI As Currency) As String
'���ܣ��ж��������Ƿ���ԭ�ۺ��ִ��޶��ķ�Χ��
'������curL=ԭ��,curR=�ּ�,curI=������
'���أ�������ڷ�Χ��,��Ϊ��ʾ��Ϣ,����Ϊ�մ�
    If (curL >= 0 And curR >= 0) Or (curL <= 0 And curR <= 0) Then
        '�����ֵ������ͬ,���þ���ֵ�ж�
        If Abs(curI) < Abs(curL) Or Abs(curI) > Abs(curR) Then
            CheckScope = "����Ľ�����ֵ���ڷ�Χ(" & Format(Abs(curL), "0.00") & "-" & Format(Abs(curR), "0.00") & ")��."
        End If
    Else
        '������Ų���ͬ,����ԭʼ��Χ�ж�
        If curI < curL Or curI > curR Then
            CheckScope = "����Ľ��ֵ���ڷ�Χ(" & Format(curL, "0.00") & "-" & Format(curR, "0.00") & ")��."
        End If
    End If
End Function


Public Function Get�շ�ִ�п���ID(ByVal lng��Ŀid As Long, ByVal intִ�п��� As Integer, _
                ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, Optional ByVal int��Χ As Integer = 2, Optional ByVal lng����ID As Long) As Long
'���ܣ���ȡ��ҩ�շ���Ŀ��ִ�п���
'������int��Χ=1.����,2-סԺ

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
    Select Case intִ�п���
        Case 0 '0-����ȷ����
            Get�շ�ִ�п���ID = UserInfo.����ID
        Case 1 '1-�������ڿ���
            Get�շ�ִ�п���ID = lng���˿���ID
        Case 2 '2-�������ڲ���
            If int��Χ = 1 Then
                Get�շ�ִ�п���ID = lng���˿���ID
            Else
                Get�շ�ִ�п���ID = lng����ID
            End If
        Case 3 '3-����Ա���ڿ���
            Get�շ�ִ�п���ID = UserInfo.����ID
        Case 4 '4-ָ������
            strSQL = "" & _
            "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            "   From �շ�ִ�п��� A,���ű� C" & _
            "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
            "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
            " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng��Ŀid, int��Χ, lng���˿���ID)
            If Not rsTmp.EOF Then
                'ȱʡȡ����Ա���ڿ���
                rsTmp.Filter = "��������ID=" & UserInfo.����ID
                If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
                If rsTmp.EOF Then rsTmp.Filter = "��������ID=0"
                If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
            End If
        Case 5 'Ժ��ִ��(Ԥ��,������δ��)
        Case 6 '�����˿���
           Get�շ�ִ�п���ID = lng��������ID
    End Select
    If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = UserInfo.����ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExecute(strNO As String, ByVal lng���ʵ�ID As Long, ByVal int��Դ As Integer) As Boolean
'���ܣ��жϷ��ö�Ӧ�Ĵ������ҩ���Ƿ��Ѿ����
'������strNO   =���õ��ݺ�
'      ���ʵ�ID=���ڱ༭�ļ��ʵ�
'      int��Դ-1����;2סԺ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(Count(ID),0) as ��Ŀ" & _
        " From " & IIf(int��Դ = 1, "������ü�¼", "סԺ���ü�¼") & _
        " Where NO=[1] And ��¼����=2 and ���ʵ�ID=[2]" & _
        " And ��¼״̬ IN(1,3) And ִ��״̬<>1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, lng���ʵ�ID)
    
    CheckExecute = (rsTmp!��Ŀ = 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
'���ܣ����ָ�������Ƿ�������Ȩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    Err = 0: On Error GoTo errH:
    If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") > 0 _
        And InStr(strPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
        Exit Function
    End If
    
    strSQL = "Select A.����,B.��Ժ����,B.״̬,X.�������" & _
        " From ������Ϣ A,������ҳ B,������� X" & _
        " Where A.����ID=B.����ID And A.����ID=X.����ID(+)" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��Ժ����) And Nvl(rsTmp!״̬, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, "��Ժ����ǿ�Ƽ���") = 0 Then
            If Nvl(rsTmp!�������, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetExamineItem(ByVal strItems As String, ByVal lngMediCareID As Long) As ADODB.Recordset
'����:����ָ��������շ���ĿҪ�������ļ�¼��
'����:strItems-�շ�ϸĿID��,����:"2369,2367,2368"
'     lngMediCareID-����,����:901
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select /*+ rule */ A.�շ�ϸĿid" & vbNewLine & _
            "From ����֧����Ŀ A ,Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B" & vbNewLine & _
            "Where A.���� = [1] And A.Ҫ������ = 1 And A.�շ�ϸĿid = B.Column_Value"
            
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngMediCareID, strItems)
    
    Set GetExamineItem = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRowByFeeItemID(ByRef ObjBillDetails As BillDetails, ByRef lngItemID As Long) As Long
'����:�����շ���ĿID�������ڵ����е��к�,������ظ���,ֻ���ص�һ��
    Dim i As Long
    
    For i = 1 To ObjBillDetails.Count
        If lngItemID = ObjBillDetails(i).�շ�ϸĿID Then
            GetRowByFeeItemID = i: Exit Function
        End If
    Next
End Function

Public Function CheckExamine(ByRef ObjBillDetails As BillDetails, ByRef rsMedAudit As ADODB.Recordset, ByRef lngMediCareID As Long) As Boolean
'����:���ݸ������շ���Ŀ���󼯺Ͳ���������Ŀ��¼�������Ӧ���շ���Ŀ�Ƿ���Ҫ����
    Dim i As Long, j As Long, strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    For i = 1 To ObjBillDetails.Count
        strTmp = strTmp & "," & ObjBillDetails(i).�շ�ϸĿID
    Next
    Set rsTmp = GetExamineItem(Mid(strTmp, 2), lngMediCareID)
    
    strTmp = ""
    For i = 1 To rsTmp.RecordCount
        rsMedAudit.Filter = "��ĿID=" & rsTmp!�շ�ϸĿID
        If rsMedAudit.RecordCount = 0 Then
            strTmp = strTmp & "," & GetRowByFeeItemID(ObjBillDetails, rsTmp!�շ�ϸĿID)
        ElseIf Not IsNull(rsMedAudit!��������) Then
            j = GetRowByFeeItemID(ObjBillDetails, rsTmp!�շ�ϸĿID)
            If ObjBillDetails(j).���� > rsMedAudit!�������� Then
                MsgBox "��" & j & "���շ���Ŀ�����γ�������׼��ʹ������" & rsMedAudit!�������� & ".", vbInformation, gstrSysName
                CheckExamine = False: Exit Function
            End If
        End If
        
        rsTmp.MoveNext
    Next
    
    If strTmp <> "" Then
        MsgBox "��" & Mid(strTmp, 2) & "���շ���ĿҪ������,��ǰ����δ����׼ʹ��!", vbInformation, gstrSysName
        CheckExamine = False: Exit Function
    End If
    CheckExamine = True
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲿����Ϣ�Ƿ���ر���
    '����:��ʾ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 13:11:01
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    If mlng���ű���ƽ������ = 0 Then
        strSQL = "Select Avg(length(����)) As ���� From ���ű�"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���ű����ƽ������")
        mlng���ű���ƽ������ = Val(Nvl(rsTemp!����))
    End If
    '���ڱ��볤�ȿ��ܹ���,�޷���ʾ���ŵ�����,����Զ���ʾ�Ͳ���ʾ����,������5ʱ,����ʾ.С��5ʱ,��ʾ
   zlIsShowDeptCode = mlng���ű���ƽ������ <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, ByVal strSearch As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮,
    Dim strIDs As String
    
    
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID = 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
    
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:

    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        rsTemp.Sort = "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    zlCommFun.PressKey vbKeyTab
    zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function




Public Function zlGetRegEventsCons(Optional strFieldName As String = "����", Optional strAliaName As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Һ���Ŀ����������
    '���:strFieldName-�������⻹�ֶ�(�缱��)
    '       strAliaName:����
    '����:
    '����:��������
    '����:���˺�
    '����:2010-12-20 16:33:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strTimeName As String
    strFieldName = IIf(strAliaName <> "", strAliaName & ".", "") & strFieldName
    strTimeName = IIf(strAliaName <> "", strAliaName & ".", "") & " �Ǽ�ʱ��"
    
    With gSysPara.Sy_Reg
        strWhere = ""
        If .bytNODaysGeneral <> 0 Or .bytNoDayseMergency <> 0 Then
            If .bytNODaysGeneral <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=0  And " & strTimeName & ">Trunc(Sysdate-" & .bytNODaysGeneral & "))"
            Else
                strWhere = strWhere & " Or  nvl(" & strFieldName & ",0)=0   "
            End If
            If .bytNoDayseMergency <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=1  And " & strTimeName & ">Trunc(Sysdate-" & .bytNoDayseMergency & "))"
            Else
                strWhere = strWhere & " Or nvl(" & strFieldName & ",0)=1  "
            End If
        End If
        If strWhere <> "" Then
            strWhere = " And  (" & Mid(strWhere, 4) & ")"
        End If
    End With
    zlGetRegEventsCons = strWhere
End Function
Public Function zlIsAllowFeeChange(lng����ID As Long, lng��ҳID As Long, _
   Optional int״̬ As Integer = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�������ñ䶯
    '���:int״̬-(-1��ʾ�����ݿ��ж�ȡ��˱�־�����ж�;>0��ʾ,ֱ�Ӹ��ݸ�״̬�����ж�)
    '����:����䶯����true,���򷵻�False
    '����:���˺�
    '����:2012-05-21 15:44:47
    '����:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    If gSysPara.byt������˷�ʽ = 0 And gSysPara.blnδ��ƽ�ֹ���� = False Then
        ''����Ǹ��
        zlIsAllowFeeChange = True: Exit Function
    End If
   
    strSQL = "" & _
    " Select Nvl(��˱�־,0) as ��˱�־,nvl(״̬,0) as ״̬" & _
    " From ������ҳ " & _
    " Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    If rsTemp.EOF Then
        MsgBox "δ�ҵ���Ӧ�Ĳ�����Ϣ,��������м�¼����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '���δ��Ʋ��˲��������
    If gSysPara.blnδ��ƽ�ֹ���� And Val(Nvl(rsTemp!״̬)) = 1 Then
        '51612
        MsgBox "����δ���(��" & lng��ҳID & "��סԺ) ,���ܶԸò��˽��м��˻����˲�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����ؼ��
    If gSysPara.byt������˷�ʽ = 0 Then zlIsAllowFeeChange = True: Exit Function
    If int״̬ < 0 Then
        int״̬ = Val(Nvl(rsTemp!��˱�־))
    End If
    '������״̬
    If int״̬ = 1 Then
        MsgBox "�����ڵ�" & lng��ҳID & "��סԺ���Ѿ���ʼ��˷���,���ܶԸò��˽��з��ñ䶯��", vbInformation, gstrSysName
        Exit Function
    End If
    If int״̬ = 2 Then
        MsgBox "�Ѿ�����˶Բ��˵�" & lng��ҳID & "��סԺ���õ����,���ܶԸò��˽��з��ñ䶯��", vbInformation, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

