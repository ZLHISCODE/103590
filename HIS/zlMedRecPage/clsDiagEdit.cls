VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmDiagEdit As frmDiagEdit
Private mclsProperty As clsProperty                   '�����࣬�������沿�ֱ�������������
' strTag=������Ϣ�����ڴ洢���ﲡ����Ƭ�ļ���·�����Ժ���չʱ����|�ָ�
Public Event Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String) '��ҳ�ر��¼�
'��ҳ�Ƿ��Ǵ�״̬��ֻ������
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'��ҳ�Ƿ�ϲ���,ֻ������
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'��Ժ��Ҫ����Ƿ����ı�,ֻ������
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:ָ����ϵ���ҳ����Ƿ�����
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitDiagEdit(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule)
'���ܣ���ʼ��
'������objMip=��Ϣ����
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f���ѡ��, lngSys, lngModule, objMip)
End Sub


Public Function ShowDiagEdit(frmParent As Object, ByVal lng��ʶID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                        ByVal int������Դ As Integer, ByVal lng��������ID As Long, ByRef str���IDs As String, _
                        ByRef str���S As String, ByVal bytSize As Byte, Optional ByVal lngҽ����ID As Long) As Boolean
'���ܣ�������ϱ༭����
'������lng����ID=����ID
'      lng����ID=סԺ:��ҳID,����Һŵ�ID
'      int������Դ=1-���2-סԺ
'      lng��������ID=�������ڿ��ң����ʹ��
'      lng��ʶID =�������ָ������뵥�ı�ʶ�����ڱ�����Ӧ�����
'      str���IDs=�����뵥��ص����ID,������ʱ���ID�Զ��ŷָ�
'      str���S =ѡ�����ϵ��ַ����������������Զ��ŷָ�
'      bytSize=0-9�����壬1-12������
'      lngҽ����ID=���뵥��Ӧ����ҽ��ID���Ѿ�����ҽ�������뵥�޸�ʱ����
'���أ� ShowDiagEdit= ��ȷ������ȡ��
'        str���S=������������ַ����������뵥ʹ��
    Dim arrTmp As Variant, i As Long
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.Module = IIf(int������Դ = 2, pסԺҽ��վ, p����ҽ��վ)
    mclsProperty.PatiType = IIf(int������Դ = 2, PF_סԺ, PF_����)
    If Not mclsProperty.CurrentForm Is mfrmDiagEdit Or mclsProperty.CurrentForm Is Nothing Then
        Set mfrmDiagEdit = New frmDiagEdit
        Set mclsProperty.CurrentForm = mfrmDiagEdit
    End If
    mclsProperty.Privs = "" '�Զ���ȡ
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.IsComfirmInfect = False
    mclsProperty.����ID = lng����ID
    mclsProperty.��ҳID = lng����ID
    mclsProperty.��Ժ����ID = lng��������ID
    mclsProperty.FontSize = IIf(bytSize = 0, 9, 12)
    mclsProperty.AplyMark = lng��ʶID
    mclsProperty.AdviceID = lngҽ����ID
    mclsProperty.DiagRowIDs = str���IDs
    mclsProperty.DiagNames = str���S
    mclsProperty.IsOK = mfrmDiagEdit.ShowMe
    Call mclsProperty.InitCacheRecInfo
    str���IDs = mclsProperty.DiagRowIDs
    str���S = mclsProperty.DiagNames
    '���ԭ�и����뵥������ϵĻ��棬�����¹��������
    If mclsProperty.IsOK Then
        Call DeleteApplyDiag(lng��ʶID)
        If str���IDs <> "" Then
            arrTmp = Split(str���IDs, ",")
            With mclsProperty.DiagConn
                For i = LBound(arrTmp) To UBound(arrTmp)
                    .AddNew
                    !���ID = Val(arrTmp(i))
                    !��ʶID = lng��ʶID
                    !���� = 0
                    .Update
                Next
            End With
        End If
    End If
    ShowDiagEdit = mclsProperty.IsOK
End Function

Public Function DeleteApplyDiag(ByVal lng��ʶID As Long) As Boolean
'���ܣ�ɾ�����뵥ʱ����������뵥���������
    With mclsProperty.DiagConn
        .Filter = "��ʶID=" & lng��ʶID
        .Sort = "���ID"
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
    End With
End Function

Public Function ConfirmInfectiousDiseases(ByVal frmParent As Object, ByVal lng��ʶID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int������Դ As Integer, ByVal lng��������ID As Long, ByVal rsInfec As ADODB.Recordset, Optional ByRef str���IDs As String, _
                        Optional ByRef str���S As String, Optional ByVal bytSize As Byte) As Boolean
'���ܣ�ȷ�ϴ�Ⱦ��
'������lng����ID=����ID
'      lng����ID=סԺ:��ҳID,����Һŵ�ID
'      lng��ʶID =�������ָ������뵥�ı�ʶ�����ڱ�����Ӧ�����
'      int������Դ=1-���2-סԺ
'      lng��������ID=�������ڿ��ң����ʹ��
'      bytSize=0-9�����壬1-12������
'      str���IDs=�����뵥��ص����ID,������ʱ���ID�Զ��ŷָ�[��ʱδʹ�ã��Ժ���չʹ�á�
'      str���S =ѡ�����ϵ��ַ����������������Զ��ŷָ�[��ʱδʹ�ã��Ժ���չʹ�á�
'���أ� ShowDiagEdit= ��ȷ������ȡ��
'        str���S=������������ַ����������뵥ʹ��[��ʱδʹ�ã��Ժ���չʹ�á�
    Dim arrTmp As Variant, i As Long
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.Module = IIf(int������Դ = 2, pסԺҽ��վ, p����ҽ��վ)
    mclsProperty.PatiType = IIf(int������Դ = 2, PF_סԺ, PF_����)
    mclsProperty.IsComfirmInfect = True
    If Not mclsProperty.CurrentForm Is mfrmDiagEdit Or mclsProperty.CurrentForm Is Nothing Then
        Set mfrmDiagEdit = New frmDiagEdit
        Set mclsProperty.CurrentForm = mfrmDiagEdit
    End If
    mclsProperty.Privs = "" '�Զ���ȡ
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.����ID = lng����ID
    mclsProperty.��ҳID = lng����ID
    mclsProperty.��Ժ����ID = lng��������ID
    mclsProperty.FontSize = IIf(bytSize = 0, 9, 12)
    mclsProperty.AplyMark = lng��ʶID
    mclsProperty.DiagRowIDs = str���IDs
    mclsProperty.DiagNames = str���S
    Do While Not rsInfec.EOF
        With mclsProperty.DiagConn
            .AddNew
            !��ʶID = lng��ʶID
            !���� = 1
            !����Ŀ¼ID = Val(rsInfec!����id & "")
            !���Ŀ¼ID = Val(rsInfec!���ID & "")
            .Update
        End With
        rsInfec.MoveNext
    Loop
    mclsProperty.IsOK = mfrmDiagEdit.ShowMe
    Call mclsProperty.InitCacheRecInfo
    str���IDs = mclsProperty.DiagRowIDs
    str���S = mclsProperty.DiagNames
    '���ԭ�и����뵥������ϵĻ��棬�����¹��������
    If mclsProperty.IsOK Then
        Call DeleteApplyDiag(lng��ʶID)
        If str���IDs <> "" Then
            arrTmp = Split(str���IDs, ",")
            With mclsProperty.DiagConn
                For i = LBound(arrTmp) To UBound(arrTmp)
                    .AddNew
                    !���ID = Val(arrTmp(i))
                    !��ʶID = lng��ʶID
                    !���� = 0
                    .Update
                Next
            End With
        End If
    End If
    ConfirmInfectiousDiseases = mclsProperty.IsOK
End Function

Public Sub Closed(ByVal blnEditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String, ByVal strTag As String)
'���ܣ������¼�
' strTag=������Ϣ�����ڴ洢���ﲡ����Ƭ�ļ���·�����Ժ���չʱ����|�ָ�
    RaiseEvent Closed(blnEditCancel, str����ID, str���ID, strTag)
End Sub

Private Sub Class_Terminate()
    If Not mfrmDiagEdit Is Nothing Then
        Unload mfrmDiagEdit
        Set mfrmDiagEdit = Nothing
    End If
    Set mclsProperty = Nothing
End Sub