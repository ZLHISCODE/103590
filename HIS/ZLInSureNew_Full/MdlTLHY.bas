Attribute VB_Name = "MdlTLHY"
Option Explicit
'���������淶:ȫ�ֱ�����g��ͷ,ģ�鼶������m��ͷ
'API��������ʾ��
'Public Declare Function BJ_Hosp_Divide3 Lib "FYFJ.dll" Alias "Hosp_Divide3" (ByVal strIn As String) As Long

'������"//TODO:�������ѵ�ʵ�ִ���"���ҵ��������㣬��Щ����㶼�Ǳ�����д�����
'-------------------------------------------------------------------------------
'��̲���˵��
'1��Ϊ���ӿڲ�������������zl9I_xxx���籱��ҽ������������Ϊ��zl9I_BJYB��ע�⣬��ģ����Ҫ����Ϊ��clsI_xxx
'2�������Ҫ��������ҽ����ص����ݣ����½�һ���û����������ǳ�֮Ϊ�м��
'3����ҽ����صĲ������ã����м����û��������������������������ӱ��ղ������ô��壬��������frmSetҽ�����ƣ��磺frmSet������
'4����������ṩ����ҽ����Ŀ�嵥������Ŀ¼�ȣ����ڱ�����Ŀѡ�����Ŀ���°�ť����д���룬��ɴ��ļ������Ľ�����·����ݸ��µ�HIS����
'5����д�������ҽ����Ŀ����Ĺ���
'6����д������������֤����
'7���������º�������̵�������룬���ҽ���ӿڵ����幦��
'8�����ݽӿ����ʣ��޸���ģ����GetCapability()��������ز�����μ�mdlInsure�е�ö�ٱ���"ҽԺҵ��"
'9��������Ҫ�޸���ģ�������������ĵ��ô���
'10��������Ҫ���ӻ��޸Ĺ��������ģ��

Public Declare Function GetMyLastError Lib "HopsInterface.dll" () As String '��ȡ���һ�δ��������
Public Declare Function FreeDllSession Lib "HopsInterface.dll" () As Long '�Ͽ��붯̬���ӿ������
Public Declare Function GetRyInfo Lib "HopsInterface.dll" _
                           (ByVal sCHRYBM As String, _
                            ByRef sXM As String, ByRef sXB As String, ByRef sSFZHM As String, _
                            ByRef sCHJTBH As String, ByRef sCSRQ As String, ByRef sHZXM As String, _
                            ByRef sHZXB As String, ByRef SHZSFZHM As String, ByRef sHKDZ As String, _
                            ByRef SYHZGXMC As String, ByRef cJTZHYE As Double) As Long '��ȡ��Ա��Ϣ
Public Declare Function SetMZFYBXData Lib "HopsInterface.dll" _
                            (ByVal sCHRYBM As String, ByVal sJZYY As String, ByVal sJZRQ As String, _
                            ByVal sBXR As String, ByVal sJZYS As String, ByVal sJZKS As String, _
                            ByVal sZDJG As String, ByVal sSFTJ As String, ByVal sSFMZMXB As String, _
                            ByVal cFSZFY As Double, ByRef cJTZHZFJE As Double, ByRef cBCZFY As Double, _
                            ByRef cXJZFJE As Double, ByRef sMZSJH As String, _
                            ByVal sLYZTDM As String, ByVal sBZDM As String, ByVal sZLKSDM As String, ByVal sJZLXDM As String, _
                            ByVal cXYF As Double, ByVal cZYF As Double, ByVal cJCF As Double, ByVal cZLF As Double, ByRef cTCJJZF As Double) As Long 'Ԥ��
Public Declare Function SaveMZFYBXData Lib "HopsInterface.dll" (ByVal sMZSJH As String) As Long '����Ԥ��
Public Declare Function StricktheBalanceMZFYBX Lib "HopsInterface.dll" (ByVal sMZSJH As String) As Long '�����������
Public Declare Function InitDllSession Lib "HopsInterface.dll" () As Long '��ʼ���붯̬���ӿ������
Public Declare Function SetZYRegister Lib "HopsInterface.dll" _
                        (ByVal sCHRYBM As String, ByVal sJZYS As String, ByVal sJZKS As String, _
                        ByVal dRYRQ As String, ByVal sZYH As String, ByRef sZYJZDH As String, _
                        ByVal sJZLXDM As String, ByVal sLYZTDM As String, ByVal sYSZJDM As String, ByVal sRYZDXX As String, ByVal sRYDJLX As String) As Long 'סԺ�Ǽ�
Public Declare Function CancelRegister Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal CHRYBM As String) As Long 'ȡ��סԺ�Ǽ�
Public Declare Function SetZYFYBXYPMX Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal NHYWBM As String, ByVal YYYWMC As String, _
                        ByVal sl As Double, ByVal je As Double, ByVal dj As Double, _
                        ByRef NBXH As Long) As Long '�ϴ�ҩƷ������ϸ
Public Declare Function ModiZYFYBXYPMX Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal NBXH As Integer) As Long 'ҩƷ������ϸ����
Public Declare Function SetZYFYBXZLMX Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal NHZLBM As String, ByVal YYZLMC As String, _
                        ByVal sl As Double, ByVal je As Double, ByVal dj As Double, _
                        ByRef NBXH As Long) As Long '�ϴ����Ʒ�����ϸ
Public Declare Function ModiZYFYBXZLMX Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal NBXH As Integer) As Long  '���Ʒ��ó���
Public Declare Function SetZYFYBXCWMX Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal NHCWBM As String, ByVal YYCWMC As String, _
                        ByVal TS As Double, ByVal je As Double, ByVal dj As Double, _
                        ByRef NBXH As Long) As Long '�ϴ���λ������ϸ
Public Declare Function ModiZYFYBXCWMX Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal NBXH As Integer) As Long  '��λ������ϸ����
Public Declare Function ZYcheckout Lib "HopsInterface.dll" _
                        (ByVal ZYJZDH As String, ByVal CHRYBM As String, ByVal dCYRQ As String, _
                        ByVal sBZDM As String, ByVal sSSMC As String, ByVal sBXR As String, _
                        ByRef cFSZFY As Double, ByRef cYPFZ As Double, ByRef cJCZLFZ As Double, _
                        ByRef cCWFZ As Double, ByRef cBCFWJE As Double, ByRef cKBYPF As Double, _
                        ByRef cKBJCZLF As Double, ByRef cKBCWF As Double, ByRef cQBXBZ As Double, _
                        ByRef cQBXSJZF As Double, ByRef cBXBL As Double, ByRef cYBJE As Double, _
                        ByRef cDSNLJBC As Double, ByRef cSJBXJE As Double, ByRef cGRZFZFY As Double, _
                        ByVal sSSMCDM As String, ByVal sCYJZLX As String, ByVal sCYZTDM As String, ByVal sZWYYDM As String, _
                        ByRef cJTZHZFJE As Double, ByRef cTCJJZF As Double) As Long 'סԺԤ��
Public Declare Function SaveZYFYBXalldata Lib "HopsInterface.dll" (ByVal ZYJZDH As String) As Long '����סԺԤ��
Public Declare Function StrickthebalanceZYFYBX Lib "HopsInterface.dll" (ByVal sZYJZDH As String) As Long 'סԺ���ó���������Ժ�Ǽ�
Public Declare Function GetCHRYBM Lib "HopsInterface.dll" _
                        (ByRef CHRYBM As String, ByVal frmY As Long, ByVal frmX As Long) As Long 'ȡ�ú�ҽ����
Public Declare Function GetBZDM Lib "HopsInterface.dll" _
                        (ByRef BZDM As String, ByRef BZMC As String, ByVal frmY As Long, _
                        ByVal frmX As Long) As Long 'ȡ�ú�ҽ���ִ���,����
Public Declare Function GetJZYY Lib "HopsInterface.dll" (ByRef YLJGDM As String, ByRef YLJGMC As String, ByVal frmX As Long, ByVal frmY As Long) As Long 'ȡҽԺ����
Public Declare Function GetSSMCDM Lib "HopsInterface.dll" (ByRef sSSMCDM As String, ByRef sSSMC As String, ByVal frmX As Long, ByVal frmY As Long) As Long 'ȡ��������
Public Declare Function ModiZYRegisterInfo Lib "HopsInterface.dll" (ByVal ZYJZDH As String, ByVal CHRYBM As String, ByVal dRYRQ As String, ByVal sZYH As String, ByVal sJZLXDM As String, ByVal sLYZTDM As String, ByVal sYSZJDM As String, ByVal sRYZDXX As String) As Long '�޸���Ժ��Ϣ

Public mҽ����ʼ�� As Boolean


Public Function ҽ����ʼ��_ͭ����ҽ() As Boolean
'>Beging ҽ����ʼ��
Dim lngReturn As Long
    If mҽ����ʼ�� = False Then
        lngReturn = InitDllSession
        If lngReturn <> 1 Then
            MsgBox "������Ϣ:" & GetMyLastError & " ��ʼ��ʧ��,���ܽ��к�ҽ����", vbInformation, "��ҽ������Ϣ"
            
            Exit Function
        Else
            ҽ����ʼ��_ͭ����ҽ = True
        End If
    End If
'>End ҽ����ʼ��

End Function


Public Function ҽ����ֹ_ͭ����ҽ() As Boolean
Dim lngReturn As Long
    If mҽ����ʼ�� = True Then
        lngReturn = FreeDllSession
        If lngReturn <> 1 Then
            MsgBox GetMyLastError, vbInformation, "ҽ��������Ϣ"
            Exit Function
        Else
            ҽ����ֹ_ͭ����ҽ = True
        End If
        mҽ����ʼ�� = False
    End If
End Function

Public Function ��ݱ�ʶ_ͭ����ҽ(Optional bytType As Byte, Optional lng����ID As Long = 0, Optional ByRef intinsure As Integer = 0) As String
    '������  ���÷�����������ò���������ҺŲ�������Ժ�Ǽǲ�������
    '����ʱ���������������س�ʱ
    '����˵���������֤�ɹ��󣬽�������Ϣ�����ظ���������
    
    Dim strReturn As String
    strReturn = frm�����֤_ͭ����ҽ.GetIdentify(bytType, lng����ID, type_ͭ����ҽ)
    'if bytType=0 then
    '   ����������Ǽǹ���
    'end if
    ��ݱ�ʶ_ͭ����ҽ = strReturn
End Function

''Public Function ҽ������_����(ByVal intInsure As Integer) As Boolean
''    'ҽ������_���� = frmSet����.��������()
''End Function

Public Function ����Һ�(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷���������ҺŲ�������
    '����ʱ�����������ҺŴ����ȷ����ťʱ
    '����˵����ͨ������ҽ���̵�����ҺŽӿڣ��ֽⱾ�η�����ϸ���õ��������������ʻ����١�ҽ��������ٵȣ�������
    'ע�������Ҫ���ù���zl_���˽����¼_Update�Բ���Ԥ����¼������������
    
End Function

Public Function ����Һų���(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷���������ҺŲ�������
    '����ʱ�����������ҺŴ���ĳ�����ťʱ
    '����˵����ͨ������ҽ���̵�����Һų����ӿڣ��������ҺŽ��������
    
End Function

Public Function �����������_ͭ����ҽ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer) As Boolean
    '������  ���÷�����������ò�������
    '����ʱ������������շѴ����Ԥ���㰴ťʱ
    '����˵����ͨ������ҽ���̵�Ԥ���㷽�����ֽⱾ�η�����ϸ���õ��������������ʻ����١�ҽ��������ٵȣ�����������������ʽ�����ڲ�����str���㷽ʽ����
    
    '����˵��
    '1������ӿ���Ҫ������÷�����ϸ�ϴ��ӿڣ���������ϸ�ϴ�
    '2����������Ԥ����ӿ�
    '3�������������涨��ʽ����
    
    '//TODO:�������ѵ�ʵ�ִ���
    'rs��ϸ��¼�����Ǳ���¼������ﴦ����ϸ
    'str���㷽ʽ�ĸ�ʽ˵����������ʽ;���;�Ƿ������޸�|....
'    str���㷽ʽ = "ҽ������;" & dblҽ������ & ";0"
'    str���㷽ʽ = str���㷽ʽ & "|���֧��;" & dbl���֧�� & ";0"

Dim �ܽ�� As Double, ״̬ As String, �������� As String, ������ As String, ҽ���� As String, ��ҽ��Ϣ
Dim R�����ʻ� As Double, Rҽ������ As Double, R�����ܷ��� As Double, R�Ը����� As Double, R��ˮ�� As String * 32
Dim rsTmp As New ADODB.Recordset, ����ҩ�� As Double, ����ҩ�� As Double, �ܼ��� As Double, �����Ʒ� As Double

On Error GoTo errHandle

    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_ͭ����ҽ = True
        Exit Function
    End If
    
    'ȡ����
    gstrSQL = "select * from �����ʻ� where ����=" & type_ͭ����ҽ & " and ����id=" & rs��ϸ("����id")
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "ȡ���տ���")
    ������ = rs��ϸ("������")
    ҽ���� = rsTmp("ҽ����")
    ��ҽ��Ϣ = Split(rsTmp("��ҽ��Ϣ"), "|")
    
    Select Case ��ҽ��Ϣ(4)
        Case "һ��"
            ״̬ = "3"
        Case "Σ"
            ״̬ = "1"
        Case "��"
            ״̬ = "2"
        Case Else
            ״̬ = "4"
    End Select
    �������� = IIf(��ҽ��Ϣ(3) = 1, "4", IIf(��ҽ��Ϣ(1) = 1, "3", "1"))
    
    '�ϴ��ܷ��ü�Ԥ��
    Do Until rs��ϸ.EOF
        If Val(rs��ϸ!����) < 0 Or Val(rs��ϸ!ʵ�ս��) < 0 Then
            MsgBox "��ҽ����֧�ָ�������", vbInformation, "�������"
            �����������_ͭ����ҽ = False
            str���㷽ʽ = ""
            Exit Function
        End If
        �ܽ�� = �ܽ�� + rs��ϸ("ʵ�ս��")
        rs��ϸ.MoveNext
    Loop
    
    If SetMZFYBXData(ҽ����, ��ҽ��Ϣ(0), CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd")), UserInfo.����, ������, "����", ��ҽ��Ϣ(5), IIf(��ҽ��Ϣ(1) = "0", "��", "��"), IIf(��ҽ��Ϣ(2) = "0", "��", "��"), �ܽ��, R�����ʻ�, R�����ܷ���, R�Ը�����, R��ˮ��, ״̬, Nvl(rsTmp!���ִ���), "", ��������, ����ҩ��, ����ҩ��, �ܼ���, �����Ʒ�, Rҽ������) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError()
        str���㷽ʽ = ""
        �����������_ͭ����ҽ = False
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        str���㷽ʽ = "�����ʻ�;" & R�����ʻ� & ";0|ҽ������;" & Rҽ������ & ";0"
        �����������_ͭ����ҽ = True
        Exit Function
    End If
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �������_ͭ����ҽ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, ByVal intinsure As Integer) As Boolean
    '������  ���÷�����������ò�������
    '����ʱ������������շѴ���Ľ��㰴ťʱ
    '����˵���������������ӿ�
    
    '����˵��
    '1�������Ҫ�ϴ���ϸ�����������ϸ�ϴ��ӿ�
    '2�������������ӿ�
    '3������ɹ����򱣴汣�ս����¼
    
    '//TODO:�������ѵ�ʵ�ִ���

Dim �ܽ�� As Double, ״̬ As String, �������� As String, ���ִ��� As String, ������ As String, ҽ���� As String, ����ID As Long, �ʻ���� As Currency, ��ҽ��Ϣ
Dim R�����ʻ� As Double, Rҽ������ As Double, R�����ܷ��� As Double, R�Ը����� As Double, R��ˮ�� As String * 32
Dim rsTmp As New ADODB.Recordset, ����ҩ�� As Double, ����ҩ�� As Double, �ܼ��� As Double, �����Ʒ� As Double


On Error GoTo errHandle
    'ȡ����
    gstrSQL = "select * from �����ʻ� where ����=" & type_ͭ����ҽ & " and ҽ����='" & strSelfNo & "'"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "ȡ���տ���")
    ����ID = rsTmp("����id")
    ҽ���� = rsTmp("ҽ����")
    ��ҽ��Ϣ = Split(rsTmp("��ҽ��Ϣ"), "|")
    �ʻ���� = Nvl(rsTmp("�ʻ����"))
    
    Select Case ��ҽ��Ϣ(4)
        Case "һ��"
            ״̬ = "3"
        Case "Σ"
            ״̬ = "1"
        Case "��"
            ״̬ = "2"
        Case Else
            ״̬ = "4"
    End Select
    
    �������� = IIf(��ҽ��Ϣ(3) = 1, "4", IIf(��ҽ��Ϣ(1) = 1, "3", "1"))
    ���ִ��� = Nvl(rsTmp!���ִ���)
    gstrSQL = "select * from ������ü�¼ where ����id=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "ȡ������Ϣ")
    ������ = rsTmp("������")
    
    '�ϴ��ܷ���,�����ܷ��ü�Ԥ��
    Do Until rsTmp.EOF
            �ܽ�� = �ܽ�� + rsTmp("ʵ�ս��")
            Select Case rsTmp!�վݷ�Ŀ
                Case "��ҩ��", "�г�ҩ��"
                    ����ҩ�� = ����ҩ�� + rsTmp!ʵ�ս��
                Case "��ҩ��"
                    ����ҩ�� = ����ҩ�� + rsTmp!ʵ�ս��
                Case "����", "�����"
                    �ܼ��� = �ܼ��� + rsTmp!ʵ�ս��
                Case Else
                    �����Ʒ� = �����Ʒ� + rsTmp!ʵ�ս��
            End Select
        rsTmp.MoveNext
    Loop
    
    If SetMZFYBXData(ҽ����, ��ҽ��Ϣ(0), CStr(Format(zlDatabase.Currentdate, "yyyy-mm-dd")), UserInfo.����, ������, "����", ��ҽ��Ϣ(5), IIf(��ҽ��Ϣ(1) = "0", "��", "��"), IIf(��ҽ��Ϣ(2) = "0", "��", "��"), �ܽ��, R�����ʻ�, R�����ܷ���, R�Ը�����, R��ˮ��, ״̬, ���ִ���, "", ��������, ����ҩ��, ����ҩ��, �ܼ���, �����Ʒ�, Rҽ������) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError()
        �������_ͭ����ҽ = False
        Exit Function
    Else
        If SaveMZFYBXData(R��ˮ��) <> 1 Then
            Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError()
            �������_ͭ����ҽ = False
            Exit Function
        Else
            gstrSQL = "zl_���ս����¼_insert(1," & _
                                                lng����ID & _
                                              "," & type_ͭ����ҽ & _
                                              "," & ����ID & _
                                              "," & Format(zlDatabase.Currentdate, "YYYY") & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0" & _
                                              ",0," & �ܽ�� & _
                                              "," & R�Ը����� & _
                                              "," & IIf(R�����ܷ��� <> 0, �ܽ�� - R�����ܷ���, R�����ܷ���) & _
                                              ",0" & _
                                              "," & Rҽ������ & _
                                              ",0" & _
                                              ",0" & _
                                              "," & R�����ʻ� & _
                                              ",'" & Trim(MidUni(R��ˮ��, 1, 32)) & "'" & _
                                              ",null" & _
                                              ",null" & _
                                              ",'Ժ��:" & ��ҽ��Ϣ(0) & "|���:" & ��ҽ��Ϣ(1) & "|����:" & ��ҽ��Ϣ(2) & "|����:" & ��ҽ��Ϣ(3) & "|״̬:" & ��ҽ��Ϣ(4) & "|���:" & ��ҽ��Ϣ(5) & "|" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss") & _
                                              "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
            �ʻ���� = �ʻ���� - R�����ʻ�
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & ����ID & "," & type_ͭ����ҽ & ",'�ʻ����','" & �ʻ���� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ����")
            �������_ͭ����ҽ = True
            Exit Function
        End If
    End If
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_ͭ����ҽ(ByVal lng����ID As Long, ByVal cur�����ʻ� As Currency, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷�����������ò�������
    '����ʱ������������շ�����������ϰ�ťʱ
    '����˵������������������Ͻӿ�

    '����˵��
    '1�����ӿڹ����ж��Ƿ��������һ�ξ�������ﵥ�ݿ�ʼ�˷�
    '2����������������Ͻӿ�
    '3�����汣�ս����¼
    
    '//TODO:�������ѵ�ʵ�ִ���

Dim rsTmp As New ADODB.Recordset
Dim lng����ID As Long

On Error GoTo errHand
'���ݴ���Ľ���id���ҳ�������id
    gstrSQL = "Select a.����id As �½���id From ����Ԥ����¼ a /*�¼�¼*/ ,����Ԥ����¼ b Where a.No = b.No And a.��¼���� = b.��¼���� And a.��¼���� = 3 And a.��¼״̬ = 2 And b.����ID = " & lng����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "ȡ����ID")
    lng����ID = rsTmp!�½���id
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & type_ͭ����ҽ & " and ����id=" & lng����ID & " and ��¼id=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "ȡ���ս����¼")
    
    If StricktheBalanceMZFYBX(rsTmp("֧��˳���")) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError()
        rsTmp.Close
        ����������_ͭ����ҽ = False
        Exit Function
    Else
        '���汣�ս����¼
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & type_ͭ����ҽ & "," & rsTmp("����id") & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & ",0,0,0,0,0,0,0,0," & _
            -1 * Nvl(rsTmp!�������ý��, 0) & "," & -1 * Nvl(rsTmp!ȫ�Ը����, 0) & "," & -1 * rsTmp!�����Ը���� & ",0," & -1 * rsTmp!ͳ�ﱨ����� & ",0,0," & _
            -1 * Nvl(rsTmp!�����ʻ�֧��, 0) & ",'" & rsTmp!֧��˳��� & "',null,null,'" & Nvl(rsTmp!��ע) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        gstrSQL = "select * from �����ʻ� where ����=" & type_ͭ����ҽ & " and ����id=" & lng����ID
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "�����")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & type_ͭ����ҽ & ",'�ʻ����','" & rsTmp!�ʻ���� + cur�����ʻ� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ����")
        ����������_ͭ����ҽ = True
        Exit Function
    End If
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_ͭ����ҽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷����ɲ�����Ժ��������
    '����ʱ���������Ժ�ǼǴ����ȷ����ťʱ
    '����˵����������Ժ�Ǽǽӿ�

    '����˵��
    '1���Ӳ�����ҳ����ȡ��Ժ���ڣ�������Ժ�Ǽ�Ҳ�ǵ��øýӿڣ���˲���ȡ��ǰ������Ϊ��Ժ�����ϴ���
    '2��������Ժ�Ǽǽӿ�
    '3��ִ����Ժ�Ǽǹ���(zl_�����ʻ�_��Ժ)�����Ĳ��˵ĵ�ǰ״̬
Dim ҽ���� As String, RסԺ����� As String * 32, S�������� As String, S��Ժ״̬ As String, Sҽ������ As String, S��Ժ��� As String, S��Ժ���� As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo errHand
    
    gstrSQL = "select * from �����ʻ� where ����=" & type_ͭ����ҽ & " and ����id=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "���ҽ����")
    ҽ���� = rsTmp!ҽ����
    
    gstrSQL = "select B.���� as ��Ժ����,A.*,C.������Ϣ from ������ҳ A,���ű� B,������ C where A.����=" & type_ͭ����ҽ & " and A.����id=" & lng����ID & " and A.��ҳid=" & lng��ҳID & " and A.��Ժ����id=B.id and a.����id=c.����id and a.��ҳid=c.��ҳid and c.�������=1"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "�鲡��ҳ�е���Ժʱ��")
    S�������� = IIf(Nvl(rsTmp("סԺĿ��")) = "����", 5, IIf(rsTmp("סԺĿ��") = "����", 2, 9))
    S��Ժ״̬ = IIf(Nvl(rsTmp("��Ժ����")) = "һ��", 3, IIf(rsTmp("��Ժ����") = "��", 2, 1))
    Sҽ������ = ""
    S��Ժ��� = Nvl(rsTmp("������Ϣ"), "��")
    S��Ժ���� = IIf(Nvl(rsTmp("��Ժ��ʽ")) = "ת��", "Zr", "Nml")
    
    '���ӿ�,�����סԺ��Ϊ����id&��ҳid
    If SetZYRegister(ҽ����, Nvl(rsTmp!סԺҽʦ, rsTmp!����ҽʦ), rsTmp!��Ժ����, CStr(Format(rsTmp!��Ժ����, "YYYY-MM-DD")), CStr(rsTmp!����ID) & "_" & CStr(rsTmp!��ҳID), RסԺ�����, S��������, S��Ժ״̬, Sҽ������, S��Ժ���, S��Ժ����) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError() & "����Ǽ�ʧ��"
        ��Ժ�Ǽ�_ͭ����ҽ = False
        Exit Function
    Else
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_ͭ����ҽ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & type_ͭ����ҽ & ",'˳���','''" & Trim(MidUni(RסԺ�����, 1, 32)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ���㵥��")
        ��Ժ�Ǽ�_ͭ����ҽ = True
        Exit Function
    End If
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽǳ���_ͭ����ҽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷����ɲ�����Ժ��������
    '����ʱ���������Ժ�ǼǴ����ȡ����ťʱ
    '����˵�������ó�����Ժ�Ǽǻ��Ժ�Ǽǽӿ�
Dim rsTmp As New ADODB.Recordset

On Error GoTo errHand

    gstrSQL = "select * from �����ʻ� where ����=" & type_ͭ����ҽ & " and ����id=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "���ҽסԺ�����")
    If CancelRegister(rsTmp!˳���, rsTmp!ҽ����) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError()
        ��Ժ�Ǽǳ���_ͭ����ҽ = False
        Exit Function
    Else
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_ͭ����ҽ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���Ĳ��˵ĵ�ǰ״̬")
        gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
        
        gstrSQL = "select * from סԺ���ü�¼ where ����id=" & lng����ID & " and ��ҳid=" & lng��ҳID
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "������˷��ü�¼�����ϴ�Ϊ0")
        
        Do Until rsTmp.EOF
            DoEvents
            gstrSQL = "zl_���˷��ü�¼_����ҽ��(" & rsTmp!ID & ",null,null,null,null,0,0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            rsTmp.MoveNext
        Loop
        
        ��Ժ�Ǽǳ���_ͭ����ҽ = True
        Exit Function
    End If
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_ͭ����ҽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷����ɲ������Ժ��������
    '����ʱ���������Ժ�����ȷ����ťʱ
    '����˵�������ó�Ժ�Ǽǽӿ�

Dim R���ִ��� As String, R�������� As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo errHand

    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_ͭ����ҽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���ճ�Ժ")
    
    ��Ժ�Ǽ�_ͭ����ҽ = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_ͭ����ҽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷����ɲ������Ժ��������
    '����ʱ�����ڳ�Ժ�������򣬵��������Ժ�˵�ʱ
    '����˵�������ó�����Ժ�Ǽǻ���Ժ�Ǽǽӿ�

    '����˵��
    '1�����ӿڹ�����м��
    '2�����ó�����Ժ�Ǽǻ���Ժ�Ǽǽӿ�
    '3��ִ����Ժ�Ǽǹ���(zl_�����ʻ�_��Ժ)�����Ĳ��˵ĵ�ǰ״̬
    
    '//TODO:�������ѵ�ʵ�ִ���
On Error GoTo errHand

    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ���Ժ")
    ��Ժ�Ǽǳ���_ͭ����ҽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_ͭ����ҽ(ByVal ����ID As Long, ByVal strҽ���� As String, ByVal intinsure As Integer) As Currency
    '������  ���÷����������շѲ�����סԺ���㲿������
    '����ʱ��������������Ҫ�˽⵱ǰҽ�����˵ĸ����ʻ����������
    '����˵�������ø����ʻ�����ѯ�ӿڻ�ֱ�Ӵӱ����ʻ�������ȡ�����ʻ����

    '����˵��
    '1�����ò�ѯ�ӿڻ�ȡ�����ʻ������±����ʻ���
    '2������ֱ�Ӵӱ����ʻ�����ȡ�����ʻ����
    
    '//TODO:�������ѵ�ʵ�ִ���
    '������� = 0
Dim rsTmp As New ADODB.Recordset
    gstrSQL = "select �ʻ���� from �����ʻ� where ����=" & type_ͭ����ҽ & " and ҽ����='" & strҽ���� & "' and ����id=" & ����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "������ʻ����")
    �������_ͭ����ҽ = rsTmp("�ʻ����")
    rsTmp.Close
End Function

Public Function סԺ����_ͭ����ҽ(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '������  ���÷�����סԺ���㲿������
    '����ʱ�������סԺ���㴰���е�ȷ����ťʱ
    '����˵������ɱ���סԺ���õ�ҽ������

    '����˵��
    '1������סԺ����ӿ�
    '2�����סԺ���㷵�صĽ�������סԺԤ���㷵�صĲ�һ�£���Ҫ����zl_���˽����¼_Update���̽�������
    
    '//TODO:�������ѵ�ʵ�ִ���

Dim sסԺ���㵥�� As String, s��ҽ���� As String, s��Ժ���� As String, s���ִ��� As String, s�������� As String
Dim R�����ܷ��� As Double, RҩƷ�ܷ��� As Double, R��������ܷ��� As Double, R��λ�ܷ��� As Double, R������Χ��� As Double
Dim R�ɱ�ҩƷ�� As Double, R�ɱ�������Ʒ� As Double, R�ɱ���λ�� As Double, R���߱�׼ As Double, Rʵ��֧������ As Double
Dim R�������� As Double, RӦ����� As Double, R��ʱ���ۼƲ������ As Double, Rʵ����� As Double, R�Ը����� As Double, R�����ʻ� As Double, Rҽ������ As Double
Dim rsTmp As New ADODB.Recordset, S�������� As String, S��Ժ���� As String, S��Ժ״̬ As String, Sת��ҽԺ���� As String, Sת��ҽԺ���� As String

On Error GoTo errHandle

    S�������� = Space(10)
    s�������� = Space(100)
    Sת��ҽԺ���� = Space(20)
    Sת��ҽԺ���� = Space(100)
    If GetSSMCDM(S��������, s��������, 150, 150) <> 1 Then
        S�������� = ""
    End If
    
    gstrSQL = "select distinct A.˳���,A.ҽ����,B.��Ժ����,A.���ִ���,A.�������� ,B.��ҳid,B.��Ժ��ʽ from �����ʻ� A,������ҳ B,����Ԥ����¼ C where C.��¼����=2 and A.����id=B.����id and B.����id=C.����id and A.����=" & type_ͭ����ҽ & " and B.��ҳid=C.��ҳid and C.����id=" & lng����ID & " and C.����id=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "������,����,��Ժ����,���ִ���,����")
    sסԺ���㵥�� = rsTmp!˳���
    s��ҽ���� = CStr(rsTmp!ҽ����)
    
    If Nvl(Trim(rsTmp!���ִ���)) = "" Or Nvl(Trim(rsTmp!��������)) = "" Then
        MsgBox "������Ϊ��ҽ���ˣ�������ѡ����", vbInformation, "������ʾ"
        סԺ����_ͭ����ҽ = False
        Exit Function
    End If
    s���ִ��� = CStr(Trim(rsTmp!���ִ���))
    s�������� = Trim(rsTmp!��������)
    s��Ժ���� = Format(rsTmp!��Ժ����, "YYYY-MM-DD")
    S��Ժ���� = IIf(rsTmp!��Ժ��ʽ = "תԺ", "Zc", "Nml")
    
    Select Case rsTmp!��Ժ��ʽ
        Case "����"
            S��Ժ״̬ = 1 '����
        Case "��ת"
            S��Ժ״̬ = 2 '��ת
        Case "תԺ"
            S��Ժ״̬ = 3 'תԺ,תԺ����Ҫ��תԺ����
            If GetJZYY(Sת��ҽԺ����, Sת��ҽԺ����, 150, 150) <> 1 Then
                MsgBox "������Ϣ" & GetMyLastError() & vbCrLf & "��Ժ��ʽΪתԺ,����Ҫ��ת��ҽԺ����,����ʧ��", vbInformation, "��ҽ������Ϣ"
                סԺ����_ͭ����ҽ = False
                Exit Function
            End If
        Case "����" '����
            S��Ժ״̬ = 4
    End Select
    
    '�Ƚӿ�Ԥ��,��ӿڽ���,��󱣴�����¼
    If ZYcheckout(sסԺ���㵥��, s��ҽ����, s��Ժ����, s���ִ���, s��������, UserInfo.����, R�����ܷ���, RҩƷ�ܷ���, R��������ܷ��� _
                        , R��λ�ܷ���, R������Χ���, R�ɱ�ҩƷ��, R�ɱ�������Ʒ�, R�ɱ���λ��, R���߱�׼, Rʵ��֧������, R�������� _
                        , RӦ�����, R��ʱ���ۼƲ������, Rʵ�����, R�Ը�����, S��������, S��Ժ����, S��Ժ״̬, Sת��ҽԺ����, R�����ʻ�, Rҽ������) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ��" & GetMyLastError() & vbLf & "         Ԥ��ʧ�ܣ����ܽ���"
        סԺ����_ͭ����ҽ = False
    ElseIf MsgBox("��������Ҫ���ʽ����鷳�����Ҫ������?" & vbLf & "��[��]����,��[��]ȡ��", vbOKCancel Or vbQuestion, "�������") = vbOK Then
        If SaveZYFYBXalldata(sסԺ���㵥��) <> 1 Then
            Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ" & GetMyLastError()
            סԺ����_ͭ����ҽ = False
        Else
            gstrSQL = "zl_���ս����¼_insert(2," & _
                                        lng����ID & _
                                        "," & type_ͭ����ҽ & _
                                        "," & lng����ID & _
                                        "," & Format(zlDatabase.Currentdate, "YYYY") & _
                                        ",0" & _
                                        ",0" & _
                                        ",0" & _
                                        "," & R��ʱ���ۼƲ������ & _
                                        "," & rsTmp!��ҳID & _
                                        "," & R���߱�׼ & _
                                        ",0," & Rʵ��֧������ & _
                                        "," & R�����ܷ��� & _
                                        "," & R�Ը����� & _
                                        "," & R�����ܷ��� - R������Χ��� & _
                                        "," & R������Χ��� & _
                                        "," & Rҽ������ & _
                                        ",0" & _
                                        ",0" & _
                                        "," & R�����ʻ� & _
                                        ",'" & sסԺ���㵥�� & "'" & _
                                        "," & rsTmp!��ҳID & _
                                        ",0,'����:" & CStr(R��������) & "|����:" & MidUni(s���ִ���, 1, 30) & "|����:" & MidUni(s��������, 1, 200) & "|" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
            סԺ����_ͭ����ҽ = True
        End If
    Else
        סԺ����_ͭ����ҽ = False
    End If
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_ͭ����ҽ(ByVal rsԤ����ϸ As ADODB.Recordset, ByVal lng����ID As Long, ByVal intinsure As Integer) As String
'������  ���÷�����סԺ���㲿������
'����ʱ�������벡����Ϣ��ѡ���˺�
'����˵������ɱ���סԺ���õ�ҽ��Ԥ����

Dim DҽԺ���� As Double, str��ʾ As String
Dim sסԺ���㵥�� As String, s��ҽ���� As String, s��Ժ���� As String, s���ִ��� As String, s�������� As String
Dim R�����ܷ��� As Double, RҩƷ�ܷ��� As Double, R��������ܷ��� As Double, R��λ�ܷ��� As Double, R������Χ��� As Double
Dim R�ɱ�ҩƷ�� As Double, R�ɱ�������Ʒ� As Double, R�ɱ���λ�� As Double, R���߱�׼ As Double, Rʵ��֧������ As Double
Dim R�������� As Double, RӦ����� As Double, R��ʱ���ۼƲ������ As Double, Rʵ����� As Double, R�Ը����� As Double, R�����ʻ� As Double, Rҽ������ As Double
Dim rs���� As New ADODB.Recordset, S�������� As String, S��Ժ���� As String, S��Ժ״̬ As String, Sת��ҽԺ���� As String, Sת��ҽԺ���� As String

On Error GoTo errHandle

'��סԺ�����
    gstrSQL = "select A.˳���,A.ҽ����,B.��Ժ����,A.���ִ���,A.��������,B.��Ժ��ʽ from �����ʻ� A,������ҳ B where A.����id=B.����id and A.����=" & type_ͭ����ҽ & " and B.��ҳid=" & rsԤ����ϸ!��ҳID & " and A.����id=" & rsԤ����ϸ!����ID
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, "������,����,��Ժ����,���ִ���,����")
    
    If rs����.EOF Or IsNull(rs����!˳���) Then
        MsgBox "�ò��˿��ܺ�ҽ������δ�ɹ�,�����²���", vbInformation, "�������"
        סԺ�������_ͭ����ҽ = ""
        Exit Function
    End If
    sסԺ���㵥�� = rs����!˳���
    s��ҽ���� = CStr(rs����!ҽ����)
    s��Ժ���� = IIf(IsNull(rs����!��Ժ����) Or rs����!��Ժ���� = "", Format(zlDatabase.Currentdate, "YYYY-MM-DD"), Format(rs����!��Ժ����, "YYYY-MM-DD"))
    
    Select Case rs����!��Ժ��ʽ
        Case "����"
            S��Ժ״̬ = 1 '����
        Case "��ת"
            S��Ժ״̬ = 2 '��ת
        Case "תԺ"
            S��Ժ״̬ = 3 'תԺ,תԺ����Ҫ��תԺ����
        Case "����" '����
            S��Ժ״̬ = 4
    End Select
    
    S��Ժ���� = "Nml" 'Ԥ��ʱΪ��ѡתԺ����,�̶�����������Nml
    'ûѡ���ֵĲ��������
    If Nvl(Trim(rs����!���ִ���)) = "" Or Nvl(Trim(rs����!��������)) = "" Then
        MsgBox "������Ϊ��ҽ���ˣ�������ѡ����", vbInformation, gstrSysName
        סԺ�������_ͭ����ҽ = ""
        Exit Function
    End If
    s���ִ��� = Trim(rs����!���ִ���)
    s�������� = Trim(rs����!��������)
    
    Do Until rsԤ����ϸ.EOF
        DҽԺ���� = DҽԺ���� + rsԤ����ϸ!���
        Call �����ϴ�_ͭ����ҽ(rsԤ����ϸ!NO, rsԤ����ϸ!��¼����, rsԤ����ϸ!��¼״̬, str��ʾ, rsԤ����ϸ!����ID, type_ͭ����ҽ)
        rsԤ����ϸ.MoveNext
    Loop
    
    '���´���Ϊ�ж��Ƿ񷵻�Ԥ����
    If ZYcheckout(sסԺ���㵥��, s��ҽ����, s��Ժ����, s���ִ���, s��������, UserInfo.����, R�����ܷ���, RҩƷ�ܷ���, R��������ܷ��� _
                    , R��λ�ܷ���, R������Χ���, R�ɱ�ҩƷ��, R�ɱ�������Ʒ�, R�ɱ���λ��, R���߱�׼, Rʵ��֧������, R�������� _
                    , RӦ�����, R��ʱ���ۼƲ������, Rʵ�����, R�Ը�����, S��������, S��Ժ����, S��Ժ״̬, Sת��ҽԺ����, R�����ʻ�, Rҽ������) <> 1 Then
        MsgBox "��Ϊ����ԭ��:" & GetMyLastError() & "Ԥ��ʧ��", vbInformation, "��ҽ������Ϣ"
        סԺ�������_ͭ����ҽ = ""
    ElseIf R�����ܷ��� = DҽԺ���� Then
        סԺ�������_ͭ����ҽ = "�����ʻ�;" & R�����ʻ� & ";0|ҽ������;" & Rҽ������ & ";0"
    Else
        If MsgBox("���ܲ��ݷ����ϴ�ʧ��,ҽԺ����:  " & DҽԺ���� & " ���ҽ���ķ���" & R�����ܷ��� & "����" & vbLf & "��[��]����,��[��]ȡ��", vbOKCancel Or vbQuestion + vbDefaultButton2, "�������") = vbOK Then
            סԺ�������_ͭ����ҽ = "�����ʻ�;" & R�����ʻ� & ";0|ҽ������;" & Rҽ������ & ";0"
        Else
            סԺ�������_ͭ����ҽ = ""
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_ͭ����ҽ(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'������  ���÷�����סԺ���㲿������
'����ʱ������ĳ�ν����������ʱ
'����˵������ɱ���סԺ���������
'���ջ�20050605ʵ�ֲ���
'���ӿڳɹ���
'1.��ȫ�����ü�¼�ϴ���־���
'2.���±����ʻ�״̬Ϊ0
'3.������ҳ.������Ϣ���������
'4.���汣�ս����¼

Dim rs���� As New ADODB.Recordset, �½���id As Long

On Error GoTo errHand

    gstrSQL = "Select a.����id As �½���id From ����Ԥ����¼ a /*�¼�¼*/ ,����Ԥ����¼ b Where a.No = b.No And a.��¼���� = 12 and b.��¼����=2 And b.����ID = " & lng����ID
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, "�������¼����id")
    �½���id = rs����!�½���id
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=" & type_ͭ����ҽ & " and ��¼id=" & lng����ID
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, "�鱣�ս����")
    
    If StrickthebalanceZYFYBX(rs����!֧��˳���) <> 1 Then
        Err.Raise 9000, "��ҽ������Ϣ", "������Ϣ:" & GetMyLastError()
        סԺ�������_ͭ����ҽ = False
    Else
        gstrSQL = "Select ID From סԺ���ü�¼ Where ��¼����=2 and ����id In (Select b.����id From ����Ԥ����¼ a /*���ʼ�¼*/,����Ԥ����¼ b/*ԭ���ʼ�¼*/ Where a.����id=" & lng����ID & " And a.����id=b.����id And a.��ҳid=b.��ҳid And b.��¼����=2)"
        Call zlDatabase.OpenRecordset(rs����, gstrSQL, "��Ҫ���ʵķ��ü�¼")
        '��ȫ�����ü�¼�ϴ���־���
        If rs����.EOF <> True Then rs����.MoveFirst
        
        Do While Not rs����.EOF
            gstrSQL = "zl_���˷��ü�¼_����ҽ��(" & rs����!ID & ",null,null,null,null,0,0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ղ��˷��ü�¼�е�ȫ��������Ϣ")
            rs����.MoveNext
        Loop
        
        gstrSQL = "select * from ���ս����¼ where ����=" & type_ͭ����ҽ & " and ����=2 and ��¼id=" & lng����ID
        Call zlDatabase.OpenRecordset(rs����, gstrSQL, "�ڱ��ս����¼�в鲡��id����ҳid")
        '���±����ʻ�״̬Ϊ1
        gstrSQL = "zl_�����ʻ�_��Ժ(" & rs����!����ID & "," & type_ͭ����ҽ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���±����ʻ��в���״̬Ϊ0")
        '������ҳ.������Ϣ���������
        gstrSQL = "zl_������ҳ_����ҽ����Ժ(" & rs����!����ID & "," & rs����!��ҳID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²�����ҳ������=null")
        '���汣�ս����¼
        gstrSQL = "zl_���ս����¼_insert(2," & _
                                                    �½���id & _
                                                    "," & type_ͭ����ҽ & _
                                                    "," & rs����!����ID & _
                                                    "," & Format(zlDatabase.Currentdate, "YYYY") & _
                                                    ",0" & _
                                                    ",0" & _
                                                    ",0" & _
                                                    "," & rs����!�ۼ�ͳ�ﱨ�� & _
                                                    "," & rs����!��ҳID & _
                                                    "," & rs����!���� & _
                                                    ",0," & rs����!ʵ������ & _
                                                    "," & -1 * rs����!�������ý�� & _
                                                    "," & -1 * rs����!ȫ�Ը���� & _
                                                    "," & -1 * (rs����!�������ý�� - rs����!����ͳ����) & _
                                                    "," & -1 * rs����!����ͳ���� & _
                                                    "," & -1 * rs����!ͳ�ﱨ����� & _
                                                    ",0" & _
                                                    ",0" & _
                                                    "," & -1 * rs����!�����ʻ�֧�� & _
                                                    ",'" & rs����!֧��˳��� & "'" & _
                                                    "," & rs����!��ҳID & _
                                                    ",0,'" & rs����!��ע & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        סԺ�������_ͭ����ҽ = True
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �����ϴ�_ͭ����ҽ(ByVal str���ݺ� As String, ByVal int���� As Integer, ByVal int״̬ As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0, Optional ByVal intinsure As Integer = 0) As Boolean
    '������  ��סԺ���ʻ�ҽ������ģ�����
    '����ʱ����סԺ���ʱ���ʱ�򱣴�󣬸��ݲ���������support...���ɲμ�Getcapability��
    '����˵������ɱ��δ�����ϸ���ϴ�

    '����˵��
    '1����ȡ�����ݵĴ�����ϸ
    '2�����ϴ���ҽ���Ĳ��˴���
    '3�����ݽӿ����ʣ�ÿ�������ϴ������ϴ��������ɹ��ϴ�����ϸ�����ϴ����
    
    '//TODO:�������ѵ�ʵ�ִ���
On Error GoTo errHand
    If int״̬ = 1 Then
    
        Dim s������ˮ�� As Long, sסԺ����� As String, �Ƿ�ɹ��ϴ� As Boolean, ��ҽ���� As String
        Dim rs������ϸ As New ADODB.Recordset
        Dim rs��ʱ As New ADODB.Recordset
        Dim rs�������ʱ� As New ADODB.Recordset
        Dim Conn�ϴ� As New ADODB.Connection
        Set Conn�ϴ� = GetNewConnection
        '���ʱ�������ȡ����idȻ��ѭ���ϴ�
        gstrSQL = "select distinct ����id from סԺ���ü�¼ where ��¼����=2 and  no='" & str���ݺ� & "'"
        Call zlDatabase.OpenRecordset(rs�������ʱ�, gstrSQL, "������ʱ��в���id")
        If rs�������ʱ�.EOF <> True Then rs�������ʱ�.MoveFirst
        
        Do While Not rs�������ʱ�.EOF
            gstrSQL = "select A.id,A.NO,A.����id,A.ʵ�ս��,A.��׼����,B.����,A.���㵥λ,A.�շ�ϸĿid,B.����,A.����*nvl(A.����,1) as ���� ,A.�շ����" & _
                        " from סԺ���ü�¼ A,�շ�ϸĿ B,������ҳ C " & _
                        "where A.no='" & str���ݺ� & "' and A.��¼����=" & int���� & " and A.��¼״̬=1 And (A.�Ƿ��ϴ�=0 or A.�Ƿ��ϴ� is null)" & _
                        " and A.����ID=C.����ID and A.��ҳID=C.��ҳID And C.����=" & type_ͭ����ҽ & " and A.�շ�ϸĿid=B.id" & _
                        " and A.����ID=" & rs�������ʱ�!����ID
            Call zlDatabase.OpenRecordset(rs������ϸ, gstrSQL, "�鱾�α�ҽ�����ʼ�¼")
            
        
            �����ϴ�_ͭ����ҽ = True
            If rs������ϸ.EOF <> True Then rs������ϸ.MoveFirst
            Do While Not rs������ϸ.EOF
                gstrSQL = "select * from  �����ʻ� where ����=" & type_ͭ����ҽ & " and ����id=" & rs������ϸ!����ID
                Call zlDatabase.OpenRecordset(rs��ʱ, gstrSQL, "��סԺ�����")
                'û�н���ҽ���ǼǵĲ��˲����ϴ��ӿ�
                If rs��ʱ.EOF = True Then
                    str��Ϣ = "����IDΪ" & rs������ϸ!����ID & "�Ĳ���û��ҽ����¼,���Ƚ���ҽ���Ǽ�"
                    Exit Function
                Else
                    sסԺ����� = rs��ʱ!˳���
                    If Nvl(rs������ϸ!ʵ�ս��) <> 0 Then '���Ϊ0�Ĳ��ϴ�
                        '���ڸ������ʷ���ֻ������ʾ,������������
                        If Val(rs������ϸ!����) < 0 Or Val(rs������ϸ!ʵ�ս��) < 0 Then
                            str��Ϣ = "��ҽ����֧�ָ�������,�뽫���ݺ�Ϊ" & rs������ϸ!NO & "����"
                            �����ϴ�_ͭ����ҽ = False
                        Else
                            gstrSQL = "select * from ����֧����Ŀ where ����=" & type_ͭ����ҽ & " and �շ�ϸĿid=" & rs������ϸ!�շ�ϸĿID
                            Call zlDatabase.OpenRecordset(rs��ʱ, gstrSQL, "����Ƿ����")
                            'û����Ĳ����ϴ�
                            If rs��ʱ.EOF Then
                                str��Ϣ = rs������ϸ!���� & "δ����,�ϴ�ʧ��"
                                �����ϴ�_ͭ����ҽ = False
                            Else
                                ��ҽ���� = rs��ʱ!��Ŀ����
                                '�����շ����ֱ��ϴ�����������ϸ
                                Select Case rs������ϸ!�շ����
                                Case "5", "6", "7"  'ҩƷ����
                                    gstrSQL = "Select c.���� From ҩƷĿ¼ a,ҩƷ��Ϣ b,ҩƷ���� c Where a.ҩ��id=b.ҩ��id And b.����=c.���� And a.����=" & rs������ϸ!����
                                    Call zlDatabase.OpenRecordset(rs��ʱ, gstrSQL, "�����")
                                    If SetZYFYBXYPMX(sסԺ�����, ��ҽ����, rs������ϸ!���� & "  |  " & rs��ʱ!���� & "  |  " & rs������ϸ!���㵥λ, rs������ϸ!����, rs������ϸ!ʵ�ս��, rs������ϸ!��׼����, s������ˮ��) <> 1 Then
                                        �Ƿ�ɹ��ϴ� = False
                                    Else
                                        �Ƿ�ɹ��ϴ� = True
                                    End If
                                Case "J"    '��λ����
                                    If SetZYFYBXCWMX(sסԺ�����, ��ҽ����, rs������ϸ!���� & "  |  " & rs������ϸ!���㵥λ, rs������ϸ!����, rs������ϸ!ʵ�ս��, rs������ϸ!��׼����, s������ˮ��) <> 1 Then
                                        �Ƿ�ɹ��ϴ� = False
                                    Else
                                        �Ƿ�ɹ��ϴ� = True
                                    End If
                                Case Else    '������
                                    If SetZYFYBXZLMX(sסԺ�����, ��ҽ����, rs������ϸ!���� & "  |  " & rs������ϸ!���㵥λ, rs������ϸ!����, rs������ϸ!ʵ�ս��, rs������ϸ!��׼����, s������ˮ��) <> 1 Then
                                        �Ƿ�ɹ��ϴ� = False
                                    Else
                                        �Ƿ�ɹ��ϴ� = True
                                    End If
                                End Select
                                '�ɹ��ϴ�����ǲ����������ˮ��
                                If �Ƿ�ɹ��ϴ� Then
                                    �����ϴ�_ͭ����ҽ = True
                                    gstrSQL = "zl_���˷��ü�¼_����ҽ��(" & rs������ϸ!ID & ",,,1,'" & ��ҽ���� & "',1,'" & s������ˮ�� & "')"
                                    Conn�ϴ�.Execute gstrSQL, , adCmdStoredProc
                                Else
                                    str��Ϣ = GetMyLastError()
                                    �����ϴ�_ͭ����ҽ = False
                                    If MsgBox("��Ϊ" & GetMyLastError() & vbLf & "����" & rs������ϸ!NO & rs������ϸ!���� & "�ϴ�ʧ��" & "     �Ƿ�����ϴ�", vbQuestion Or vbOKCancel, "��ҽ������Ϣ") <> vbOK Then Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
                rs������ϸ.MoveNext
            Loop
            rs�������ʱ�.MoveNext
        Loop
    Else
        �����ϴ�_ͭ����ҽ = ��������_ͭ����ҽ(str���ݺ�, int����, int״̬, str��Ϣ, lng����ID, type_ͭ����ҽ)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ��������_ͭ����ҽ(ByVal str���ݺ� As String, ByVal int���� As Integer, ByVal int״̬ As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0, Optional ByVal intinsure As Integer = 0) As Boolean

Dim sסԺ����� As String, �Ƿ�ɹ����� As Boolean
Dim rs������ϸ As New ADODB.Recordset
Dim rs��ʱ As New ADODB.Recordset
Dim rs�������ʱ� As New ADODB.Recordset
Dim Conn���� As New ADODB.Connection
Set Conn���� = GetNewConnection
On Error GoTo errHand:

    If int״̬ = 2 Then
        gstrSQL = "select distinct ����id from סԺ���ü�¼ where ��¼����=2 and no='" & str���ݺ� & "'"
        Call zlDatabase.OpenRecordset(rs�������ʱ�, gstrSQL, "������ʱ��в���id")
        If rs�������ʱ�.EOF <> True Then rs�������ʱ�.MoveFirst
        Do While Not rs�������ʱ�.EOF
        '���Ҫ�����ĺ�ҽ��ˮ��
           gstrSQL = "Select a.Id," & _
                             "a.����id," & _
                             "a.�շ����," & _
                             "b.ʵ�ս��," & _
                             "b.���ձ���," & _
                             "b.ժҪ As ������ˮ��" & _
                    " From סԺ���ü�¼ a/*�¼�¼*/, סԺ���ü�¼ b/*ԭ��¼*/" & _
                    " Where b.No = '" & str���ݺ� & "'" & _
                            " And b.��¼���� = " & int���� & _
                            " And b.��¼״̬ = 3" & _
                            " And b.�Ƿ��ϴ� = 1" & _
                            " And a.No = b.No And a.��¼���� = b.��¼����" & _
                            " And a.��� = b.��� And a.��¼״̬ =" & int״̬ & _
                            " And (a.�Ƿ��ϴ� is null or a.�Ƿ��ϴ�=0)" & _
                            " and b.����id=" & rs�������ʱ�!����ID
            Call zlDatabase.OpenRecordset(rs������ϸ, gstrSQL, "��ҽ��������ˮ��")
            
        
            If rs������ϸ.EOF <> True Then rs������ϸ.MoveFirst
            
            ��������_ͭ����ҽ = True
            Do While Not rs������ϸ.EOF
            
            gstrSQL = "select * from  �����ʻ� where ����=" & type_ͭ����ҽ & " and ����id=" & rs������ϸ!����ID
            Call zlDatabase.OpenRecordset(rs��ʱ, gstrSQL, "��סԺ�����")
            If rs��ʱ.EOF = True Then
                str��Ϣ = "����IDΪ" & rs������ϸ!����ID & "�Ĳ���û��ҽ����¼,���Ƚ���ҽ���Ǽ�"
            Else
                sסԺ����� = rs��ʱ!˳���
            rs��ʱ.Close
                If Nvl(rs������ϸ!ʵ�ս��) <> 0 Then
                    '�����շ����ֱ���ýӿڳ���������ϸ
                    Select Case rs������ϸ!�շ����
                        Case "5", "6", "7"  'ҩƷ����
                            If ModiZYFYBXYPMX(sסԺ�����, CInt(rs������ϸ!������ˮ��)) <> 1 Then
                                �Ƿ�ɹ����� = False
                            Else
                                �Ƿ�ɹ����� = True
                            End If
                        Case "J"    '��λ����
                            If ModiZYFYBXCWMX(sסԺ�����, CInt(rs������ϸ!������ˮ��)) <> 1 Then
                                �Ƿ�ɹ����� = False
                            Else
                                �Ƿ�ɹ����� = True
                            End If
                        Case Else    '������
                            If ModiZYFYBXZLMX(sסԺ�����, CInt(rs������ϸ!������ˮ��)) <> 1 Then
                                �Ƿ�ɹ����� = False
                            Else
                                �Ƿ�ɹ����� = True
                            End If
                    End Select
                    '�ɹ��������ϴ���Ǹ�Ϊ1�����������ˮ��
                    If �Ƿ�ɹ����� Then
                        gstrSQL = "zl_���˷��ü�¼_����ҽ��(" & rs������ϸ!ID & ",,,1,'" & rs������ϸ!���ձ��� & "',1,'" & rs������ϸ!������ˮ�� & "')"
                        Conn����.Execute gstrSQL, , adCmdStoredProc
                    Else
                        str��Ϣ = GetMyLastError()
                        ��������_ͭ����ҽ = False
                        If MsgBox("��Ϊ" & GetMyLastError() & vbLf & "����" & str���ݺ� & "��" & rs������ϸ!���� & "�ϴ�ʧ��" & "       �Ƿ�����ϴ�", vbQuestion Or vbOKCancel, "��ҽ������Ϣ") <> vbOK Then Exit Function
                    End If
                End If
                End If
                �Ƿ�ɹ����� = False
                rs������ϸ.MoveNext
            Loop
        rs�������ʱ�.MoveNext
        Loop
    End If
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function


Public Function ҽ����������_ͭ����ҽ(cap���� As ҽԺҵ��) As Boolean
    '
    Select Case cap����
        Case support����Ԥ��, _
             support�����˷�, _
             support������봫����ϸ, _
             support�����ϴ�, _
             support������ɺ��ϴ�, _
             support���������ϴ�, _
             supportҽ���ϴ�, _
             support������Ժ, _
             supportδ�����Ժ, _
             support����ʹ�ø����ʻ�, _
             support��Ժ��������Ժ, _
             support����¼��������, _
             support������Ժ, _
             support��Ժ���˽�������
            ҽ����������_ͭ����ҽ = True
    End Select

End Function

Public Function ����ѡ��_ͭ����ҽ(lng����ID As Long, intinsure As Integer) As Boolean
Dim R����id As String, R�������� As String
    R����id = Space(12)
    R�������� = Space(100)
    If GetBZDM(R����id, R��������, 150, 150) <> 1 Then
        MsgBox GetMyLastError() & vbCrLf & "���ָ���ʧ��", vbInformation, "��ҽ������Ϣ"
        ����ѡ��_ͭ����ҽ = False
    Else
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'���ִ���','''" & Trim(MidUni(R����id, 1, 20)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���没�ִ����ڱ����ʻ���")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'��������','''" & Trim(MidUni(R��������, 1, 200)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���没�������ڱ����ʻ���")
        ����ѡ��_ͭ����ҽ = True
    End If
End Function
