VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsRelas"
Attribute VB_Ext_KEY = "Member0" ,"clsRelas"
'�ֲ����������漯��
Private mCol As Collection

Public Function Add(objRelas As clsRelas, Optional sKey As String) As clsRelas
    '�����¶���
    Dim objNewMember As clsRelas
    Set objNewMember = objRelas
    mCol.Add objNewMember, sKey
    '�����Ѵ����Ķ���
    
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsRelas
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mCol.Remove vntIndexKey
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

Public Sub RelasCheck(ByVal strKey As String, ByVal blnChecked As Boolean)
    '���Ȩ��֮��Ĺ�ϵ����������Ӧ�Ĵ���
    
    Dim objRelas As clsRelas
    Dim int���� As Integer, int��ϵ As Integer, int������ As Integer, bln������֮һ��ѡ�� As Boolean
    int���� = Me.Item(strKey).����
    int��ϵ = Me.Item(strKey).��ϵ
    int������ = Me.Item(strKey).������
    Me.Item(strKey).Checked = blnChecked
    
    If int��ϵ = 1 Then
        For Each objRelas In Me
            If objRelas.���� = int���� And objRelas.Key <> strKey Then
                If blnChecked Then
                    Me.Item(objRelas.Key).Checked = False
                End If
            End If
        Next
    ElseIf int��ϵ = 2 Then
        If int������ = 1 Then
            '������������, ѡ�е�����£��ӹ��ܲ���������
            If Not blnChecked Then
                'δѡ�е�����£���������������Ƿ���ѡ�У������δѡ�У���ȡ���ӹ���
                bln������֮һ��ѡ�� = False
                
                '�ȴ��������ϵΪOR�����
                For Each objRelas In Me
                    If objRelas.���� = int���� And objRelas.������ = 1 Then
                        If objRelas.�����ܹ�ϵ = 0 Then
                            If objRelas.Checked Then bln������֮һ��ѡ�� = True
                        End If
                    End If
                Next
                
                '�ٴ��������ϵΪAND�����
                For Each objRelas In Me
                    If objRelas.���� = int���� And objRelas.������ = 1 Then
                        If objRelas.�����ܹ�ϵ = 1 Then
                            If objRelas.Checked And bln������֮һ��ѡ�� Then
                                bln������֮һ��ѡ�� = True
                            Else
                                bln������֮һ��ѡ�� = False
                            End If
                        End If
                    End If
                Next
                
                If Not bln������֮һ��ѡ�� Then
                    For Each objRelas In Me
                        If objRelas.���� = int���� And objRelas.������ = 0 Then
                            Me.Item(objRelas.Key).Checked = False
                        End If
                    Next
                End If
            End If
        Else
            '��������������
            If blnChecked Then
                '������ѡ�У�����Ϊѡ��
                bln������֮һ��ѡ�� = False
                For Each objRelas In Me
                    If objRelas.���� = int���� And objRelas.������ = 1 Then
                        If objRelas.�����ܹ�ϵ = 0 Then
                            If objRelas.Checked Then bln������֮һ��ѡ�� = True
                        End If
                    End If
                Next
                
                '�ٴ��������ϵΪAND�����
                For Each objRelas In Me
                    If objRelas.���� = int���� And objRelas.������ = 1 Then
                        If objRelas.�����ܹ�ϵ = 1 Then
                            If objRelas.Checked And bln������֮һ��ѡ�� Then
                                bln������֮һ��ѡ�� = True
                            Else
                                bln������֮һ��ѡ�� = False
                            End If
                        End If
                    End If
                Next
                
                If Not bln������֮һ��ѡ�� Then
                    Me.Item(strKey).Checked = False
                End If
            End If
        End If
    Else
        For Each objRelas In Me
            If objRelas.���� = int���� And objRelas.Key <> strKey Then
                Me.Item(objRelas.Key).Checked = blnChecked
            End If
        Next
    End If

End Sub