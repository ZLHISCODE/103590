VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ClinicPlanOffice 
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   ScaleHeight     =   5415
   ScaleWidth      =   7725
   Begin MSComctlLib.ListView lvwDoctorRoom 
      Height          =   2925
      Left            =   2130
      TabIndex        =   5
      Top             =   1140
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   5159
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.OptionButton opt���﷽ʽ 
      Caption         =   "������"
      Height          =   300
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Top             =   50
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton opt���﷽ʽ 
      Caption         =   "ָ������"
      Height          =   300
      Index           =   1
      Left            =   1710
      TabIndex        =   2
      Top             =   50
      Width           =   1035
   End
   Begin VB.OptionButton opt���﷽ʽ 
      Caption         =   "��̬����"
      Height          =   300
      Index           =   2
      Left            =   2775
      TabIndex        =   3
      Top             =   50
      Width           =   1035
   End
   Begin VB.OptionButton opt���﷽ʽ 
      Caption         =   "ƽ������"
      Height          =   300
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   50
      Width           =   1035
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H80000000&
      Height          =   3975
      Left            =   330
      Top             =   540
      Width           =   6225
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "���﷽ʽ"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "ClinicPlanOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'���Ա���:
Dim m_EditMode As gRegistPlanEditMode
Dim m_ҽ������ As String
Dim m_IsDataChanged As Boolean

'ȱʡ����ֵ:
Const m_def_EditMode = 0
Const m_def_ҽ������ = ""
Const m_def_IsDataChanged = False

Private mblnNotClick As Boolean
Private mobj�������Ҽ� As �������Ҽ�
Private mobj���з������� As �������Ҽ�
'�¼�����:
Event DataIsChanged()


Public Function LoadData(ByVal obj�������Ҽ� As �������Ҽ�, Optional ByVal obj���з������� As �������Ҽ�, _
    Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�������
    '���:
    '       obj�������Ҽ� - �������Ҽ�
    '       obj���з������� - ���з������Ҽ� ,������ʾ�鿴
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Err = 0: On Error GoTo Errhand:
    Set mobj�������Ҽ� = obj�������Ҽ�
    If mobj�������Ҽ� Is Nothing Then Set mobj�������Ҽ� = New �������Ҽ�
    Set mobj���з������� = obj���з�������
    
    m_IsDataChanged = blnChanged
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetEnabled���﷽ʽ(ByVal strҽ������ As String)
    If strҽ������ = "" Then
        opt���﷽ʽ(2).Enabled = m_EditMode = ED_RegistPlan_Edit
        opt���﷽ʽ(2).Tag = ""
        opt���﷽ʽ(3).Enabled = m_EditMode = ED_RegistPlan_Edit
        opt���﷽ʽ(3).Tag = ""
    Else
        opt���﷽ʽ(2).Enabled = False
        opt���﷽ʽ(2).Tag = "1"
        opt���﷽ʽ(3).Enabled = False
        opt���﷽ʽ(3).Tag = "1"
        If opt���﷽ʽ(2).Value Or opt���﷽ʽ(3).Value Then
            opt���﷽ʽ(0).Value = True
        End If
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objRoom As ��������, objCheckedRoom As ��������
    Dim intCol As Integer, intRow As Integer
    Dim ObjItem As ListItem
    
    Err = 0: On Error GoTo Errhand:
    If mobj�������Ҽ� Is Nothing Then Set mobj�������Ҽ� = New �������Ҽ�
    
    mblnNotClick = True
    opt���﷽ʽ(mobj�������Ҽ�.���﷽ʽ).Value = True
    ҽ������ = mobj�������Ҽ�.ҽ������
    mblnNotClick = False
    lvwDoctorRoom.Enabled = m_EditMode = ED_RegistPlan_Edit And mobj�������Ҽ�.���﷽ʽ <> 0
    
    With lvwDoctorRoom
        .ListItems.Clear
        .Refresh
        If mobj���з������� Is Nothing Then
            'ֻ������ѡ������
            For Each objCheckedRoom In mobj�������Ҽ�
                Set ObjItem = .ListItems.Add(, "K" & objCheckedRoom.����ID, objCheckedRoom.��������)
                ObjItem.SubItems(1) = objCheckedRoom.����ID
                ObjItem.Checked = True
            Next
        Else
            For Each objRoom In mobj���з�������
                Set ObjItem = .ListItems.Add(, "K" & objRoom.����ID, objRoom.��������)
                ObjItem.SubItems(1) = objRoom.����ID
                '������ѡ������
                For Each objCheckedRoom In mobj�������Ҽ�
                    If objRoom.����ID = objCheckedRoom.����ID Then
                        ObjItem.Checked = True: Exit For
                    End If
                Next
            Next
        End If
    End With
    lvwDoctorRoom.BackColor = lvwDoctorRoom.BackColor
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get�������Ҽ�() As �������Ҽ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������Ҽ�
    '����:
    '����:��ȡ�ɹ�������true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim objRoom As ��������, obj�������Ҽ� As New �������Ҽ�
    
    Err = 0: On Error GoTo Errhand:
    '����δ�ı䣬ֱ�ӷ���ԭ���ϵĸ���
    If m_IsDataChanged = False Then
        Set Get�������Ҽ� = mobj�������Ҽ�.Clone
        Exit Function
    End If
    
    '�����Ѹı䣬���¹��켯�϶���
    With obj�������Ҽ�
        .���﷽ʽ = GetSelectedIndex(opt���﷽ʽ)
        .ҽ������ = ҽ������
        .�Ƿ��޸� = True
    End With
    For i = 1 To lvwDoctorRoom.ListItems.Count
        If lvwDoctorRoom.ListItems(i).Checked Then
            Set objRoom = New ��������
            With objRoom
                .����ID = lvwDoctorRoom.ListItems(i).SubItems(1)
                .�������� = lvwDoctorRoom.ListItems(i).Text
            End With
            obj�������Ҽ�.AddItem objRoom, "K" & objRoom.����ID
        End If
    Next
    Set Get�������Ҽ� = obj�������Ҽ�
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub lvwDoctorRoom_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        'ָ������ֻ��ѡһ��
        If opt���﷽ʽ(1).Value Then Call ClearAllGridChecked
        Item.Checked = True
    End If
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub lvwDoctorRoom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub opt���﷽ʽ_GotFocus(index As Integer)
    opt���﷽ʽ(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub opt���﷽ʽ_LostFocus(index As Integer)
     opt���﷽ʽ(index).BackColor = Me.BackColor
End Sub

 
Private Sub UserControl_Initialize()
    Call InitFace
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_IsDataChanged = m_def_IsDataChanged
    m_ҽ������ = m_def_ҽ������
    m_EditMode = m_def_EditMode
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_ҽ������ = PropBag.ReadProperty("ҽ������", m_def_ҽ������)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With shpBack
        .Left = 0
        .Top = opt���﷽ʽ(0).Top + opt���﷽ʽ(0).Height + 60
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top
    End With
    With lvwDoctorRoom
        .Left = 10
        .Top = shpBack.Top + 10
        .Width = ScaleWidth - .Left - 20
        .Height = ScaleHeight - .Top - 10
    End With
End Sub

Public Property Get Get�����������Ҽ�() As �������Ҽ�
    Set Get�����������Ҽ� = Get�������Ҽ�
End Property

Public Function IsValied() As Boolean
    '�������
    Dim intCount As Integer, i As Long, j As Long
    
    Err = 0: On Error GoTo errHandler
    '����δ�ı䲻���
    If m_IsDataChanged = False Or m_EditMode <> ED_RegistPlan_Edit Then IsValied = True: Exit Function
    
    '�����ж�
    If opt���﷽ʽ(0).Value = False Then
        '������ʱ�ż��
        For i = 1 To lvwDoctorRoom.ListItems.Count
            If lvwDoctorRoom.ListItems(i).Checked Then
                intCount = intCount + 1
            End If
        Next
        If opt���﷽ʽ(1).Value Then
            'ָ����������ֻ��ѡ��һ��
            If intCount = 0 Then
                MsgBox "ָ������ʱ����ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                If lvwDoctorRoom.Visible And lvwDoctorRoom.Enabled Then lvwDoctorRoom.SetFocus
                Exit Function
            ElseIf intCount > 1 Then
                MsgBox "ָ������ʱֻ��ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                If lvwDoctorRoom.Visible And lvwDoctorRoom.Enabled Then lvwDoctorRoom.SetFocus
                Exit Function
            End If
        Else
            If intCount < 2 Then
                MsgBox "��̬�����ƽ������ʱ����Ҫѡ��������Ӧ���������ң�", vbInformation, gstrSysName
                If lvwDoctorRoom.Visible And lvwDoctorRoom.Enabled Then lvwDoctorRoom.SetFocus
                Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub opt���﷽ʽ_Click(index As Integer)
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    If mblnNotClick Then Exit Sub
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    lvwDoctorRoom.Enabled = m_EditMode = ED_RegistPlan_Edit And index <> 0
    lvwDoctorRoom.BackColor = lvwDoctorRoom.BackColor
    Call ClearAllGridChecked(index)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ClearAllGridChecked(Optional ByVal byt���﷽ʽ As Byte)
    '���ܣ����ѡ����Ŀ
    Dim i As Integer, j As Integer
    Dim intSelectedCount As Integer
    
    Err = 0: On Error GoTo errHandler
    For i = 1 To lvwDoctorRoom.ListItems.Count
        Select Case byt���﷽ʽ
        Case 0 '������
            lvwDoctorRoom.ListItems(i).Checked = False
        Case 1 'ָ������
            If intSelectedCount = 1 Then lvwDoctorRoom.ListItems(i).Checked = False
            If lvwDoctorRoom.ListItems(i).Checked Then intSelectedCount = 1
        Case 2 '��̬����
        Case 3 'ƽ������
        End Select
    Next
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub opt���﷽ʽ_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub InitFace()
    Err = 0: On Error GoTo errHandler
    With lvwDoctorRoom
        .Checkboxes = True
        .FullRowSelect = True
        .GridLines = False
        .HideSelection = True
        .LabelEdit = lvwManual
        .MultiSelect = True
        .View = lvwList
        .TextBackground = lvwTransparent
        
        '�����
        .ColumnHeaders.Add , "K_����", "��������", 3500
        .ColumnHeaders.Add , "K_ID", "����ID", 0
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub UserControl_Show()
    lvwDoctorRoom.View = lvwReport
    lvwDoctorRoom.View = lvwList
End Sub

Private Sub UserControl_Terminate()
    Set mobj�������Ҽ� = Nothing
    Set mobj���з������� = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("ҽ������", m_ҽ������, m_def_ҽ������)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Dim i As Integer
    For i = opt���﷽ʽ.LBound To opt���﷽ʽ.UBound
        opt���﷽ʽ(i).BackColor = New_BackColor
    Next
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,1,1,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=13,0,0,
Public Property Get ҽ������() As String
    ҽ������ = m_ҽ������
End Property

Public Property Let ҽ������(ByVal New_ҽ������ As String)
    m_ҽ������ = New_ҽ������
    PropertyChanged "ҽ������"
    
    m_ҽ������ = New_ҽ������
    SetEnabled���﷽ʽ m_ҽ������
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    Dim i As Integer
    For i = opt���﷽ʽ.LBound To opt���﷽ʽ.UBound
        'opt���﷽ʽ(i).Tag = "1"��ʾ�����޸�״̬
        opt���﷽ʽ(i).Enabled = m_EditMode = ED_RegistPlan_Edit And opt���﷽ʽ(i).Tag = ""
    Next
    lvwDoctorRoom.Enabled = m_EditMode = ED_RegistPlan_Edit And opt���﷽ʽ(0).Value = False
End Property

