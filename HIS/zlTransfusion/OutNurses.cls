VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutNurses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"OutNurse"
Attribute VB_Ext_KEY = "Member0" ,"OutNurse"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection


Public Function getOutNurse(ByVal lngDeptID As Long) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strName As String
    On Error GoTo hErr
    
    strSQL = "Select /*+Rule */ Distinct D.���� " & vbNewLine & _
            "From ��Ա�� D, ��Ա����˵�� B, ������Ա A" & vbNewLine & _
            "Where A.��Աid = D.ID And A.��Աid = B.��Աid And (D.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.����ʱ�� Is Null) And B.��Ա���� = '��ʿ' And A.����id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select Distinct D.���� " & vbNewLine & _
            "From ��Ա�� D, ��Ա����˵�� C, ������Ա B," & vbNewLine & _
            "     (Select A.ID From ���ű� A Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id) A" & vbNewLine & _
            "Where B.��Աid = D.ID And A.ID = B.����id And B.��Աid = C.��Աid And (D.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or D.����ʱ�� Is Null) And C.��Ա���� = '��ʿ'"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "outNurses.getOutNurse", lngDeptID)
    Call ClearOutNurse
    Do Until rsTmp.EOF
        strName = zlCommFun.NVL(rsTmp.Fields("����"))
        Add strName, strName
        rsTmp.MoveNext
    Loop
    Exit Function
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Add(���� As String, Optional sKey As String) As OutNurse
    '�����¶���
    Dim objNewMember As OutNurse
    Set objNewMember = New OutNurse


    '���ô��뷽��������
    objNewMember.���� = ����
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As OutNurse
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Sub ClearOutNurse()
    Dim lngIndex As Long
    
    For lngIndex = 1 To Me.Count
        Remove 1
    Next
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub
