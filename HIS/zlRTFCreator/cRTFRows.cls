VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFRows"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFRow"
Attribute VB_Ext_KEY = "Member0" ,"cRTFRow"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function LastRow() As cRTFRow
Attribute LastRow.VB_Description = "��ȡ������ӵ��ж���"
    Set LastRow = mCol(mCol.Count)
End Function

Public Function Add(Optional Row As Long, Optional RowBand As Long, _
Optional sKey As String, Optional IsLastRow As Boolean = False, _
Optional CellSpace As Long = 108, _
Optional Width As Long = 0, Optional WidthUnit As CellLengthUnitEnum = clt3_�, _
Optional WidthB As Long = 0, Optional WidthBUnit As CellLengthUnitEnum = clt3_�, _
Optional WidthA As Long = 0, Optional WidthAUnit As CellLengthUnitEnum = clt3_�, _
Optional PatType As PatTypeEnum = ptt00_��, Optional PatForeColor As ColorEnum = clr00_Ĭ��, _
Optional PatBackColor As ColorEnum = clr00_Ĭ��, Optional PatPercentage As Long, _
Optional AutoFit As Boolean = True, Optional IsHeadRow As Boolean = False, _
Optional Keep As Boolean = False, Optional KeepFollow As Boolean = False, _
Optional Alignment As RowAlignTypeEnum = ral0_Ĭ�϶���, _
Optional Height As Long = 0, Optional PadLeft As Long = 0, _
Optional PadRight As Long = 0, Optional PadTop As Long = 0, _
Optional PadBottom As Long = 0, Optional Nested As Boolean = False, _
Optional BorderLeft As Boolean = True, Optional BorderRight As Boolean = True, _
Optional BorderTop As Boolean = True, Optional BorderBottom As Boolean = True, _
Optional BorderH As Boolean = True, Optional BorderV As Boolean = True, _
Optional BorderStyle As BorderStyleEnum = cbs00_Ĭ�ϱ߿���ʽ, _
Optional BorderColor As ColorEnum = clr00_Ĭ��, Optional BorderWidth As Long = 15) As Long
Attribute Add.VB_Description = "����һ�С�"
    '�����¶���
    Dim objNewMember As cRTFRow
    Set objNewMember = New cRTFRow


    '���ô��뷽��������
    objNewMember.Row = Row
    objNewMember.RowBand = RowBand
    objNewMember.IsLastRow = IsLastRow
    objNewMember.CellSpace = CellSpace
    objNewMember.Width = Width
    objNewMember.WidthUnit = WidthUnit
    objNewMember.WidthB = WidthB
    objNewMember.WidthBUnit = WidthBUnit
    objNewMember.WidthA = WidthA
    objNewMember.WidthAUnit = WidthAUnit
    objNewMember.PatType = PatType
    objNewMember.PatForeColor = PatForeColor
    objNewMember.PatBackColor = PatBackColor
    objNewMember.PatPercentage = PatPercentage
    objNewMember.AutoFit = AutoFit
    objNewMember.IsHeadRow = IsHeadRow
    objNewMember.Keep = Keep
    objNewMember.KeepFollow = KeepFollow
    objNewMember.Alignment = Alignment
    objNewMember.Height = Height
    objNewMember.PadLeft = PadLeft
    objNewMember.PadRight = PadRight
    objNewMember.PadTop = PadTop
    objNewMember.PadBottom = PadBottom
    objNewMember.Nested = Nested
    objNewMember.BorderLeft = BorderLeft
    objNewMember.BorderRight = BorderRight
    objNewMember.BorderTop = BorderTop
    objNewMember.BorderBottom = BorderBottom
    objNewMember.BorderH = BorderH
    objNewMember.BorderV = BorderV
    objNewMember.BorderStyle = BorderStyle
    objNewMember.BorderColor = BorderColor
    objNewMember.BorderWidth = BorderWidth
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Add = mCol.Count
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFRow
Attribute Item.VB_Description = "��ȡһ���ж���"
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_Description = "��������"
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "ɾ��ָ������Ԫ�ض���"
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
