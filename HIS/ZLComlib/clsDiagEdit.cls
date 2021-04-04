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

Public Function ShowDiagEdit(frmParent As Object, ByVal lng��ʶID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, _
                        ByVal int������Դ As Integer, ByVal lng��������ID As Long, ByVal str������ As String, ByRef str���IDs As String, _
                        ByRef str���S As String, ByVal bytSize As Byte, Optional ByVal lngҽ����ID As Long) As Boolean
'���ܣ�������ϱ༭����
'������lng����ID=����ID
'      lng����ID=סԺ:��ҳID,����Һŵ�ID
'      int������Դ=1-���2-סԺ
'      lng��������ID=�������ڿ��ң����ʹ��
'      lng��ʶID =�������ָ������뵥�ı�ʶ�����ڱ�����Ӧ�����
'      str������=����Ա��������ϵǼ���
'      str���IDs=�����뵥��ص����ID,������ʱ���ID�Զ��ŷָ�
'      str���S =ѡ�����ϵ��ַ����������������Զ��ŷָ�
'      bytSize=0-9�����壬1-12������
'      lngҽ����ID=���뵥��Ӧ����ҽ��ID���Ѿ�����ҽ�������뵥�޸�ʱ����
'���أ� ShowDiagEdit= ��ȷ������ȡ��
'        str���S=������������ַ����������뵥ʹ��
    Dim blnOK As Boolean
    Dim i As Long
    Dim arrTmp As Variant
    
    blnOK = frmDiagEdit.ShowMe(frmParent, lng��ʶID, lng����ID, lng����ID, int������Դ, lng��������ID, str������, str���IDs, str���S, bytSize, lngҽ����ID)
    '���ԭ�и����뵥������ϵĻ��棬�����¹��������
    If blnOK Then
        Call DeleteApplyDiag(lng��ʶID)
        If str���IDs <> "" Then
            arrTmp = Split(str���IDs, ",")
            With grsDiagConn
                For i = LBound(arrTmp) To UBound(arrTmp)
                    .AddNew
                    !���ID = Val(arrTmp(i))
                    !��ʶID = lng��ʶID
                    .Update
                Next
            End With
        End If
    End If
    ShowDiagEdit = blnOK
End Function

Private Sub Class_Initialize()
    Set grsDiagConn = New ADODB.Recordset
    '��¼���뵥�������
    With grsDiagConn
        .Fields.Append "���ID", adInteger
        .Fields.Append "��ʶID", adInteger
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Public Function DeleteApplyDiag(ByVal lng��ʶID As Long) As Boolean
'���ܣ�ɾ�����뵥ʱ����������뵥���������
    With grsDiagConn
        .Filter = "��ʶID=" & lng��ʶID
        .Sort = "���ID"
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
    End With
End Function

Private Sub Class_Terminate()
    Set grsDiagConn = Nothing
End Sub