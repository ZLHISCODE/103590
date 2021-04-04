VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'���ò˵����������API���Ѵ����򵥵ĵ����˵���
'�ڹ���������һ���࣬����ΪcPopupMenu

Option Explicit
'

Public Event MenuClick(ByVal Key As Long, ByVal Caption As String)
Public Event MenuBeforeShow(ByRef Cancel As Boolean)
Public Event MenuAfterClick()

Private mblnCreated As Boolean                      '�Ƿ��Ѿ������˵�
Private mlngMainMenu As Long                        '�����˵�ID
Private mrsMenu As New ADODB.Recordset

Public Function Add(ByVal Key As Long, ByVal Caption As String, Optional ByVal ParentKey As Long = 0, Optional ByVal Style As Byte = 1, Optional ByVal Enabled As Boolean = True, Optional ByVal Index As Long = 0, Optional Checked As Boolean = False) As Boolean
    With mrsMenu
        .AddNew
        .Fields("����").Value = Caption
        .Fields("����").Value = Style
        .Fields("����").Value = Enabled
        .Fields("��ѡ").Value = Checked
        .Fields("���").Value = Index
        .Fields("Key").Value = Key
        .Fields("�ϼ�Key").Value = ParentKey
    End With
End Function

Public Property Get ItemValue(vntIndexKey As Variant, vntField As String) As Variant
    '���ü����е�һ��Ԫ��ʱʹ�á�
    mrsMenu.Filter = ""
    mrsMenu.Filter = "Key=" & vntIndexKey
    ItemValue = mrsMenu.Fields(vntField).Value
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    mrsMenu.Filter = ""
    Count = mrsMenu.RecordCount
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    mrsMenu.Filter = ""
    mrsMenu.Filter = "Key=" & vntIndexKey
    mrsMenu.Delete adAffectCurrent
End Sub

Private Sub CopyRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset)
    Dim lngLoop As Long
    Dim lngCol As Long
    
    '����rsTo�ṹ
    If rsTo.State = adStateOpen Then rsTo.Close
    With rsTo
        For lngCol = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields.Append rsFrom.Fields(lngCol).Name, rsFrom.Fields(lngCol).Type, rsFrom.Fields(lngCol).DefinedSize
        Next
        rsTo.Open
    End With
    
    '��������
    rsFrom.Filter = ""
    If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    For lngLoop = 1 To rsFrom.RecordCount
        rsTo.AddNew
        For lngCol = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields(lngCol).Value = rsFrom.Fields(lngCol).Value
        Next
        rsFrom.MoveNext
    Next
End Sub

Private Function CreateSubMenu(Optional ByVal strUpKey As Long = 0) As Long
    Dim lngLoop As Long
    Dim lngMenuHand As Long
    Dim lngSubMenuHand As Long
    Dim rs As New ADODB.Recordset
    Dim lngFlags As Long
    Dim lngMenu As Long
    
    '���Ƽ�¼��
    Call CopyRecord(mrsMenu, rs)
        
    rs.Filter = ""
    rs.Filter = "�ϼ�Key=" & strUpKey
    If rs.RecordCount = 0 Then Exit Function
    
    rs.MoveFirst
    
    lngMenuHand = CreatePopupMenu()
    If lngMenuHand <= 0 Then Exit Function
    
    For lngLoop = 1 To rs.RecordCount
        Select Case Val(rs.Fields("����").Value)
        Case 2
            lngFlags = MF_SEPARATOR
        Case Else
            lngFlags = MF_STRING
            If rs.Fields("����").Value = False Then lngFlags = lngFlags + MF_GRAYED
            If rs.Fields("��ѡ").Value Then lngFlags = lngFlags + MF_CHECKED
            
            
        End Select
        
        lngMenu = AppendMenu(lngMenuHand, lngFlags, Val(rs.Fields("Key").Value), rs.Fields("����").Value)
        
        '����Ƿ����Ӳ˵�
        lngSubMenuHand = 0
        lngSubMenuHand = CreateSubMenu(rs.Fields("Key").Value)
        
        '�޸��Ӳ˵���������ϵ
        If lngSubMenuHand > 0 Then Call ModifyMenu(lngMenuHand, lngLoop, lngFlags + MF_POPUP, lngSubMenuHand, CStr(rs.Fields("����").Value))
                
        rs.MoveNext
    Next
    
    CreateSubMenu = lngMenuHand
End Function

Public Function CreatePopMenu() As Boolean
    '����:���������˵�
            
    If mlngMainMenu > 0 Then
        Call DestroyMenu(mlngMainMenu)
        mlngMainMenu = 0
        mblnCreated = False
    End If

    mlngMainMenu = CreateSubMenu
    
    mblnCreated = True
    
    CreatePopMenu = True
End Function

Public Sub ShowPopupMenuByCursor()
    '����:�ڵ�ǰλ����ʾ�����˵�
    Dim p As POINTAPI
    
    Call GetCursorPos(p)
    Call ShowPopupMenu(p.X * 15, p.Y * 15)

End Sub

Public Sub ShowPopupMenu(ByVal X As Single, ByVal Y As Single)
    '����:��ָ��λ����ʾ�����˵�
    
    Dim lngMenuPostion As Long
    Dim strMenuCaption As String
    Dim blnCancel As Boolean
    
    RaiseEvent MenuBeforeShow(blnCancel)
    If blnCancel Then Exit Sub
    
    If mblnCreated = False Then
        '���û�д���,���ȴ����˵�
        Call CreatePopMenu
    End If
    
    If mlngMainMenu <= 0 Then Exit Sub
    
    lngMenuPostion = TrackPopupMenu(mlngMainMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, X / 15, Y / 15, 0, GetForegroundWindow(), 0)
    
    RaiseEvent MenuAfterClick
    
    If lngMenuPostion <= 0 Then Exit Sub
    
    strMenuCaption = Space(255)
    Call GetMenuString(mlngMainMenu, lngMenuPostion, strMenuCaption, Len(strMenuCaption), MF_BYCOMMAND)
    
    RaiseEvent MenuClick(lngMenuPostion, strMenuCaption)
    
End Sub

Private Sub Class_Initialize()
    With mrsMenu
        .Fields.Append "����", adVarChar, 30
        .Fields.Append "����", adTinyInt
        .Fields.Append "����", adBoolean
        .Fields.Append "��ѡ", adBoolean
        .Fields.Append "���", adInteger
        .Fields.Append "Key", adBigInt
        .Fields.Append "�ϼ�Key", adBigInt
        .Open
    End With
End Sub

Private Sub Class_Terminate()
    If mlngMainMenu > 0 Then Call DestroyMenu(mlngMainMenu)
    
    If mrsMenu.State = adStateOpen Then mrsMenu.Close
    Set mrsMenu = Nothing
End Sub