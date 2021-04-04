VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_YUYAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'˵�����������ӿڣ�YUYAMA�����꣩��
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mcnThird As ADODB.Connection

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mcnThird = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'˵���������ӿ�Ϊ�״̬
    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mcnThird Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "����YUYAMA������ʧ��", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "��ʼ����YUYAMA������", 1
        mcnThird.Open mobjBase.URL
        mobjBase.mobjLog.Add "�ɹ�����YUYAMA������", 1
    Else
        Set mcnThird = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Add "����YUYAMA�������쳣", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Private Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)
End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim objDB As Object
    
    'Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    
    '�����ʹ�û���ļ�¼�������뽫ǰ��������Σ����������д�µĴ���
    'TODO��
    
    On Error GoTo hErr
    
    strInfo = Split(strData, "|")(1)
    
    Select Case intType
    Case 21
        If TypeName(mobjBase.mobjComLib) = "clsComLib" Then
            'ZLHIS����������ҩƷ�Զ����豸��������
            Set objDB = mobjBase.mobjComLib.zlDatabase
        ElseIf TypeName(mobjBase.mobjComLib) = "clsPublic" Then
            '֧����
            Set objDB = mobjBase.mobjComLib
        Else
            mobjBase.mobjLog.Add "�������Ͳ���ȷ����ֹ��ȡ��׼����", 1
            mobjBase.mobjLog.Add "" & intType, 1, 1
            mobjBase.mobjLog.Save
            Exit Function
        End If
        
        strSQL = "Select /*+cardinality(z,10)*/ Distinct Zl_Fun_Drug_Machine(c.�ⷿid, g.ҩƷ����, c.Id) �ӿڱ��, c.�ⷿid, I1.���� As ��ҩҩ��," & vbNewLine & _
                "    Decode(Nvl(o.��־, 0), 11, 1, 12, 1, 0, 1, 0) �Ƿ���, b.��ʶ�� סԺ��," & vbNewLine & _
                "    c.���ܷ�ҩ�� As ��ҩ��, Sysdate As ����ʱ��, c.�Է�����id As ��ҩ����id, I2.���� As ��ҩ����, j.����id, j.���� As ��������," & vbNewLine & _
                "    j.�Ա� As �����Ա�, j.���� As ��������, j.���￨��, b.���˲���id As ����id, I3.���� As ����, j.��ǰ���� As ����," & vbNewLine & _
                "    k.���ȼ�, a.����ҽ��, a.����ʱ��, a.Id As ҽ��id, g.ҩƷ����, c.ҩƷid, e.���� As ҩƷ����, f.���� As ��Ʒ����, c.�÷�," & vbNewLine & _
                "    c.����, h.���㵥λ As ������λ, d.סԺ��λ, a.ҽ����Ч, m.Ҫ��ʱ�� As ִ��ʱ��, a.ҽ������, g.��������, a.������־," & vbNewLine & _
                "    Decode(m.Ҫ��ʱ��, Null, Round(c.ʵ������ / d.סԺ��װ, 2), Round((a.�������� / d.����ϵ��) / d.סԺ��װ, 2)) As סԺ��λ����," & vbNewLine & _
                "    d.����ϵ��, a.У��ʱ�� ��ʿУ��ʱ��, c.������ �շ�������, c.����� �շ������, c.ID �շ�ID ,m.ҽ��id," & vbNewLine & _
                "    replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','') As ʱ�����," & vbNewLine & _
                "    Case When replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','')<1200 And replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','')>=0000 Then   '����'" & vbNewLine & _
                "         When replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','')<1800 And replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','')>=1200 Then '����'" & vbNewLine & _
                "         When replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','')<0000 And replace(to_char(m.Ҫ��ʱ��,'hh24-mi'),'-','')>=1800 Then '����'" & vbNewLine & _
                "    End  As  ����ʱ��,to_char(m.Ҫ��ʱ��,'YYYY-MM-DD') As ��ҩʱ�� " & vbNewLine & _
                "From ����ҽ����¼ A, סԺ���ü�¼ B, ҩƷ�շ���¼ C, ҩƷ��� D, �շ���ĿĿ¼ E, �շ���Ŀ���� F, ҩƷ���� G, ������ĿĿ¼ H," & vbNewLine & _
                "    ���ű� I1, ���ű� I2, ���ű� I3, ������Ϣ J, ���� K, ����ҽ������ L, ҽ��ִ��ʱ�� M, ����ҽ����¼ N," & vbNewLine & _
                "    ҩƷ�շ�סԺ��־ O, Table(f_Num2list([1], ';')) Z " & vbNewLine & _
                "Where a.Id = b.ҽ����� And b.Id = c.����id And c.ҩƷid = d.ҩƷid And c.ҩƷid = e.Id And e.Id = f.�շ�ϸĿid(+) " & vbNewLine & _
                "    And d.ҩ��id = g.ҩ��id And g.ҩ��id = h.Id And c.�ⷿid = I1.Id And c.�Է�����id = I2.Id And b.���˲���id = I3.Id " & vbNewLine & _
                "    And b.����id = j.����id And j.���� = k.����(+) And a.���id = n.Id And n.Id = m.ҽ��id " & vbNewLine & _
                "    And b.ҽ����� = l.ҽ��id(+) And b.No = l.No And b.ִ�в���id = l.ִ�в���id " & vbNewLine & _
                "    And l.���ͺ� = m.���ͺ�(+) And c.Id = o.�շ�id(+) And f.����(+) = 3 And f.����(+) = 1 " & vbNewLine & _
                "    And c.Id = z.Column_Value " & vbNewLine & _
                "Order By c.�ⷿid, c.�Է�����id, j.����id"
                
        mobjBase.mobjLog.Add strSQL, 1, 1
        Set rsResult = objDB.OpenSQLRecord(strSQL, "��ȡ��ʼ��ҩ����", strInfo)
        mobjBase.mobjLog.Add "��ȡ��ʼ��ҩ�������", 1
        
        Set clsINF_GetData = rsResult
    Case Else
        'ʹ�û��������
        Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    End Select
    Exit Function
    
hErr:
    blnFinish = False
    mobjBase.mobjLog.Add "��ȡ��׼����ʧ��", 1
    mobjBase.mobjLog.Add Err.Number & "��" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_KEY(ByVal RHS As String)
    mobjBase.KEY = RHS
End Property

Private Property Get clsINF_KEY() As String
    clsINF_KEY = mobjBase.KEY
End Property

Private Property Set clsINF_mobjComLib(ByVal RHS As Object)
    Set mobjBase.mobjComLib = RHS
End Property

Private Property Get clsINF_mobjComLib() As Object
    Set clsINF_mobjComLib = mobjBase.mobjComLib
End Property

Private Property Set clsINF_mobjLog(ByVal RHS As clsLog)
    Set mobjBase.mobjLog = RHS
End Property

Private Property Get clsINF_mobjLog() As clsLog
    Set clsINF_mobjLog = mobjBase.mobjLog
End Property

Private Property Let clsINF_mstrNOs(ByVal RHS As String)
    mobjBase.mstrNOs = RHS
End Property

Private Property Get clsINF_mstrNOs() As String
    clsINF_mstrNOs = mobjBase.mstrNOs
End Property

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, Optional strWinInfo As String) As Boolean
'���ܡ����������ء�˵������μ�clsINF�����SendData����

    Dim strSQL As String
    Dim colProc As New Collection
    Dim intSN As Integer, intMacNo As Integer
    Dim lng����id As Long

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "��ʼ��������", 1
    
    Select Case intType
    Case 21
        rsData.Filter = "�Ƿ���=1"
    Case Else
        mobjBase.mobjLog.Add "���ܲ�֧��", 1
        GoTo makExit
    End Select
    
    If rsData.RecordCount <= 0 Then
        mobjBase.mobjLog.Add "�����ݿɴ���", 1
makExit:
        mobjBase.mobjLog.Save
        clsINF_SendData = True
        Exit Function
    End If
    
    '���Ի�������ѡ������ţ�
    If mdlDefine.gtypYUYAMA.BottomLevel Then
        Call frmINF_YUYAMA_MacNo.ShowMe(intMacNo)
    Else
        intMacNo = mdlDefine.gtypYUYAMA.MacNO
    End If
    
    With rsData
        intSN = 1
        Do While .EOF = False
            '�ô洢���̵����ƹ̶�
            Select Case intType
            Case Val("21-��ҩ")
                '��Ŵ���
                If lng����id <> mdlDrugMachine.NVL(!����id) Then
                    intSN = 1
                End If
                
                '�洢����
                strSQL = "ZL_ҩƷ�豸�ӿ�_YYM_INS("
                '������
                strSQL = strSQL & "To_Char(Sysdate, 'MMDDHH') || Replace(To_Char(" & mdlDrugMachine.NVL(!����id) & ", '9999999999999'), ' ', '0')"
                '���
                strSQL = strSQL & ", " & intSN
                'С����
                strSQL = strSQL & ", 1"
                '������ţ�0�����У�1-n��ָ��������Ű�ҩ��
                strSQL = strSQL & ", " & intMacNo
                '����״̬��0��δ������1��������2���쳣��3��ȡ��������-1�������У���
                strSQL = strSQL & ", 0"
                '����ID
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!����id, 0) & "'"
                '��������
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!��������) & "'"
                '������סԺ���֣�1�����2��סԺ������3��סԺ������
                strSQL = strSQL & ", " & IIf(mdlDrugMachine.NVL(!ҽ����Ч, 0) = 0, "'2'", "'3'")
                '������ұ��룩��סԺ���������룩
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!����id, 0) & "'"
                '����������ƣ���סԺ���������ƣ�
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!����, 0) & "'"
                '���봦������
                strSQL = strSQL & ", To_Date('" & Format(mdlDrugMachine.NVL(!����ʱ��, "2000/1/1"), "yyyy-mm-dd hh:mm:ss") & "', 'yyyy-mm-dd hh24:mi:ss')"
                '�������0����������1-9���������ɻ�������
                strSQL = strSQL & ", " & mdlDrugMachine.NVL(!������־, 0)
                'ҩƷ����
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!ҩƷid, 0) & "'"
                'ҩƷ����
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!ҩƷ����) & "'"
                '��ҩ��λ
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!סԺ��λ) & "'"
                '���������
                strSQL = strSQL & ", 1"     '1��1����¼
                '�÷�
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!�÷�) & "'"
                '����ʱ��
                strSQL = strSQL & ", To_Date('" & Format(mdlDrugMachine.NVL(!ִ��ʱ��, "2000/1/1"), "yyyy-mm-dd hh:mm:ss") & "', 'yyyy-mm-dd hh24:mi:ss')"
                'HIS�ķ�ҩʱ��
                strSQL = strSQL & ", Sysdate"
                strSQL = strSQL & ")"
                
            Case Else
                strSQL = ""
            End Select
            
            If strSQL <> "" Then
                mdlDrugMachine.AddArray colProc, strSQL
            End If
            
            lng����id = mdlDrugMachine.NVL(!����id)
            
            .MoveNext
            intSN = intSN + 1
        Loop
    End With
    
    '�����������洢����
    On Error GoTo makFail
    Call mdlDrugMachine.ExecuteProcedureBeach(colProc, "", mcnThird, mobjBase.mobjLog)
    On Error GoTo hErr
    
    '������ݴ��ͳɹ�
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, True)
    
    mobjBase.mobjLog.Add "��ɴ�������", 1
    mobjBase.mobjLog.Save
    
    'ִ�л���Ĵ���
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
makFail:
    '������ݴ���ʧ��
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, False)
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "��������ʧ��", 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
    '
End Property

Private Property Get clsINF_Token() As String
    '
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
    
    Set mcnThird = New ADODB.Connection
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property
