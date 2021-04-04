VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl ClinicPlanDetailPages 
   Appearance      =   0  'Flat
   BackColor       =   &H80000014&
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   ScaleHeight     =   6120
   ScaleWidth      =   11835
   Begin VB.PictureBox picTimeWork 
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   0
      Left            =   2490
      ScaleHeight     =   4110
      ScaleWidth      =   5700
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   5700
      Begin zl9RegEvent.ClinicPlanDetail ClinicDetail 
         Height          =   2925
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   570
         Width           =   5205
         _ExtentX        =   12938
         _ExtentY        =   8176
      End
   End
   Begin XtremeSuiteControls.TabControl tbPageTimeWork 
      Height          =   930
      Left            =   570
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1680
      _Version        =   589884
      _ExtentX        =   2963
      _ExtentY        =   1640
      _StockProps     =   64
   End
   Begin VB.Shape shpLine 
      BorderColor     =   &H80000003&
      Height          =   630
      Left            =   120
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "ClinicPlanDetailPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ȱʡ����ֵ:
Const m_def_ҽ������ = ""
Const m_def_BackStyle = 0
Const m_def_����Ƶ�� = 5
Const m_def_EditMode = 0
Const m_def_ForeColor = 0
Const m_def_BorderStyle = 0
'���Ա���:
Dim m_ҽ������ As String
Dim m_BackStyle As Integer
Dim m_����Ƶ�� As Integer
Dim m_EditMode As gRegistPlanEditMode
Dim m_ForeColor As Long
Dim m_Font As Font
Dim m_BorderStyle As Integer
'�¼�����:
Event DataIsChanged(index As Integer)
Event Click()
Attribute Click.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ťʱ������"
Event DblClick()
Attribute DblClick.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ť���ٴΰ��²��ͷ���갴ťʱ������"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "���û���ӵ�н���Ķ����ϰ��������ʱ������"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "���û����º��ͷ� ANSI ��ʱ������"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "���û���ӵ�н���Ķ������ͷż�ʱ������"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "���û���ӵ�н���Ķ����ϰ�����갴ťʱ������"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "���û��ƶ����ʱ������"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mobj�����¼�� As �����¼��
Private mobj���к�����λ As ������λ���Ƽ�
Private mobj������������ As �������Ҽ�
Private mblnNotClick As Boolean
Private mstrPreTabPage As String '��һ��ѡ��ҳ��
Private mblnLoaded As Boolean
Private mblnShowFirstPage As Boolean '�Ƿ�ȱʡ��ʾ��һҳ��������ʾ���һҳ

Public Function LoadData(ByVal obj�����¼�� As �����¼��, _
    Optional ByVal obj������������ As �������Ҽ�, _
    Optional ByVal obj���к�����λ As ������λ���Ƽ�, _
    Optional ByVal blnShowFirstPage As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:obj�����Դ-�����Դ��Ϣ
    '     obj������������-������Ч����������
    '     obj���к�����λ-���к�����λ
    '     blnShowFirstPage-�Ƿ�ȱʡ��ʾ��һҳ��������ʾ���һҳ
    '����:���سɹ�, ����true,���򷵻�False
    '����:���˺�
    '����:2016-01-11 14:26:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj�����¼�� = obj�����¼��
    mblnShowFirstPage = blnShowFirstPage
    Set mobj������������ = obj������������: Set mobj���к�����λ = obj���к�����λ
    LoadData = InitPageAndData   '����ҳ�漰����
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetClinicRecord(ByVal obj�����¼�� As �����¼��, ByVal strʱ��� As String) As �����¼
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ�������ȡ��Ӧ�ĳ����¼��
    '���:obj�����¼��-�����¼��
    '     strʱ���-ʱ���
    '����:�����¼����
    '����:���˺�
    '����:2016-03-24 15:37:50
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼ As �����¼
    If obj�����¼�� Is Nothing Then Exit Function
    
    On Error GoTo errHandle
    For Each obj�����¼ In obj�����¼��
        If obj�����¼.ʱ��� = strʱ��� Then
            Set GetClinicRecord = obj�����¼.Clone: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InitPageAndData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2016-01-11 14:23:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim objWork As �ϰ�ʱ��, lngRow As Long
    Dim objPlan As �����¼, obj�����¼ As �����¼
    Dim intPageCount As Integer, intSelectedPageIndex As Integer
    
    Err = 0: On Error GoTo Errhand:
    Call LockWindowUpdate(UserControl.Hwnd)
    
    tbPageTimeWork.RemoveAll
    mstrPreTabPage = "" '��ǵ�ǰΪδѡ���κ�ҳǩ
    
    'ȱʡ����5��ʱ�οؼ�
'    If ClinicDetail.Count < 5 Then
'        For i = 1 To 4
'            Load picTimeWork(i): picTimeWork(i).Visible = True
'            Load ClinicDetail(i): ClinicDetail(i).Visible = True
'            Set ClinicDetail(i).Container = picTimeWork(i)
'        Next
'    End If
    intPageCount = ClinicDetail.Count
    
    If Not mobj�����¼�� Is Nothing Then
        If mobj�����¼��.Count > 0 Then
            lngRow = 0
            For Each objPlan In mobj�����¼��
                If lngRow > intPageCount - 1 Then
                    Load picTimeWork(lngRow): picTimeWork(lngRow).Visible = True
                    Load ClinicDetail(lngRow): ClinicDetail(lngRow).Visible = True
                    Set ClinicDetail(lngRow).Container = picTimeWork(lngRow)
                End If
                Set ObjItem = tbPageTimeWork.InsertItem(lngRow + 1, objPlan.ʱ���, picTimeWork(lngRow).Hwnd, 0)
                ClinicDetail(0).EditMode = ED_RegistPlan_View
                lngRow = lngRow + 1
            Next
        Else
            '�����һ��ҳǩ������
            ClinicDetail(0).LoadData Nothing, Nothing
            ClinicDetail(0).EditMode = ED_RegistPlan_View
        End If
    End If
    If tbPageTimeWork.ItemCount = 0 Then
        lngRow = 0
        Set ObjItem = tbPageTimeWork.InsertItem(lngRow + 1, "���ϰ�ʱ��", picTimeWork(lngRow).Hwnd, 0)
        ClinicDetail(lngRow).EditMode = ED_RegistPlan_View:
    End If
    Call LockWindowUpdate(0)
    
    With tbPageTimeWork
        If mblnShowFirstPage Then
            intSelectedPageIndex = 0
        Else
            intSelectedPageIndex = tbPageTimeWork.ItemCount - 1
        End If
        .Enabled = False
        .Item(intSelectedPageIndex).Selected = True
        '�ֶ�����SelectedChanged�¼�
        Call tbPageTimeWork_SelectedChanged(.Item(intSelectedPageIndex))
        .Enabled = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Color = xtpTabColorVisualStudio
    End With
    mblnLoaded = True
    InitPageAndData = True
    Exit Function
Errhand:
    tbPageTimeWork.Visible = True
    mblnNotClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "����/���ö���ı߿���ʽ��"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    shpLine.Visible = New_BorderStyle <> 0
    UserControl_Resize
End Property

Private Sub ClinicDetail_DataIsChanged(index As Integer)
    RaiseEvent DataIsChanged(index)
End Sub

'����ҳǩ���ƻ�ȡҳǩ����
Public Property Get ItemIndex(ByVal Caption As String) As Integer
    Dim i As Integer
    
    ItemIndex = -1
    For i = 0 To tbPageTimeWork.ItemCount - 1
        If tbPageTimeWork.Item(i).Caption = Caption Then
            ItemIndex = i: Exit For
        End If
    Next
End Property

Private Sub picTimeWork_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picTimeWork(index)
        ClinicDetail(index).Left = .ScaleLeft
        ClinicDetail(index).Top = .ScaleTop
        ClinicDetail(index).Height = .ScaleHeight - 10
        ClinicDetail(index).Width = .ScaleWidth - 10
    End With
End Sub
 
Private Sub tbPageTimeWork_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim obj�����¼ As �����¼
    
    If mstrPreTabPage = Item.Caption Then Exit Sub
    
    mstrPreTabPage = Item.Caption
    If Val(Item.Tag) = 1 Then Exit Sub
    
    Item.Tag = "1"
    '��������
    If mobj�����¼�� Is Nothing Then Exit Sub
    If mobj�����¼��.Exits("K" & Item.Caption) = False Then Exit Sub
    Set obj�����¼ = mobj�����¼��("K" & Item.Caption) 'GetClinicRecord(mobj�����¼��, Item.Caption)
    If obj�����¼ Is Nothing Then Exit Sub
    
    ClinicDetail(Item.index).LoadData obj�����¼, mobj���к�����λ, mobj������������
    ClinicDetail(Item.index).EditMode = m_EditMode
    If obj�����¼.�Ƿ�̶� Then ClinicDetail(Item.index).EditMode = ED_RegistPlan_View
End Sub

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_BorderStyle = m_def_BorderStyle
    m_EditMode = m_def_EditMode
    m_BackStyle = m_def_BackStyle
    m_����Ƶ�� = m_def_����Ƶ��
    Call Set����Ƶ��(m_����Ƶ��)
    m_ҽ������ = m_def_ҽ������
    Call Setҽ������(m_ҽ������)
    mblnLoaded = False
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    
    shpLine.Visible = m_BorderStyle <> 0
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
    
    Call ReSetPageEditMode
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_����Ƶ�� = PropBag.ReadProperty("����Ƶ��", m_def_����Ƶ��)
    Call Set����Ƶ��(m_����Ƶ��)
    m_ҽ������ = PropBag.ReadProperty("ҽ������", m_def_ҽ������)
    Call Setҽ������(m_ҽ������)
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With shpLine
        .Top = ScaleTop
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    With tbPageTimeWork
        .Top = IIf(shpLine.Visible, 10, 0)
        .Left = IIf(shpLine.Visible, 10, 0)
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top * 2
    End With
End Sub
 
Private Sub UserControl_Show()
    If mblnLoaded Then Exit Sub
    Call InitPageAndData
End Sub

Private Sub UserControl_Terminate()
    Set mobj�����¼�� = Nothing
    Set mobj���к�����λ = Nothing
    Set mobj������������ = Nothing
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("����Ƶ��", m_����Ƶ��, m_def_����Ƶ��)
    Call PropBag.WriteProperty("ҽ������", m_ҽ������, m_def_ҽ������)
End Sub
 
 
Public Property Get Get�����¼��() As �����¼��
    Dim obj�����¼�� As New �����¼��, obj�����¼ As New �����¼
    Dim intPage As Integer
    
    On Error GoTo Errhand
    If mobj�����¼�� Is Nothing Then Exit Property
    If mobj�����¼��.�������� <> "" Then
        obj�����¼��.�������� = mobj�����¼��.��������
        For intPage = 0 To tbPageTimeWork.ItemCount - 1
            If tbPageTimeWork(intPage).Caption <> "���ϰ�ʱ��" And tbPageTimeWork(intPage).Caption <> "" Then
                If Val(tbPageTimeWork.Item(intPage).Tag) = 1 Then
                    Set obj�����¼ = ClinicDetail(intPage).Get�����¼
                    If obj�����¼.�Ƿ��޸� Then obj�����¼��.�Ƿ��޸� = True
                Else
                    'δ���ص�����
                    Set obj�����¼ = GetClinicRecord(mobj�����¼��, tbPageTimeWork.Item(intPage).Caption)
                End If
                obj�����¼��.AddItem obj�����¼, "K" & obj�����¼.ʱ���
            End If
        Next
    End If
    Set Get�����¼�� = obj�����¼��
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Function IsValied() As Boolean
    '�������
    Dim intPage As Integer
    
    Err = 0: On Error GoTo errHandler
    For intPage = 0 To tbPageTimeWork.ItemCount - 1
        If ClinicDetail(intPage).IsValied() = False Then
            tbPageTimeWork.Enabled = False
            tbPageTimeWork(intPage).Selected = True
            tbPageTimeWork.Enabled = True
            Exit Function
        End If
    Next
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=24,0,0,0
Public Property Get EditMode(ByVal index As Integer) As gRegistPlanEditMode
    If index < 0 Or index > tbPageTimeWork.ItemCount - 1 Then Exit Property
    EditMode = ClinicDetail(index).EditMode
End Property

Public Property Let EditMode(Optional ByVal index As Integer = -1, ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    If index = -1 Then
        '��������ҳǩ
        For index = 0 To tbPageTimeWork.ItemCount - 1
            If tbPageTimeWork(index).Caption = "���ϰ�ʱ��" Then
                '�����ÿ���״̬���Ҳ��ܸ�
                ClinicDetail(index).EditMode = ED_RegistPlan_View
            ElseIf Not mobj�����¼�� Is Nothing Then
                If index < mobj�����¼��.Count Then
                    ClinicDetail(index).EditMode = m_EditMode
                    If mobj�����¼��(index + 1).�Ƿ�̶� Then ClinicDetail(index).EditMode = ED_RegistPlan_View
                End If
            Else
                ClinicDetail(index).EditMode = m_EditMode
            End If
        Next
        Exit Property
    End If
    
    If index < 0 Or index > tbPageTimeWork.ItemCount - 1 Then Exit Property
    If Not mobj�����¼�� Is Nothing Then
        If index < mobj�����¼��.Count Then
            ClinicDetail(index).EditMode = m_EditMode
            If mobj�����¼��(index + 1).�Ƿ�̶� Then ClinicDetail(index).EditMode = ED_RegistPlan_View
        End If
    Else
        ClinicDetail(index).EditMode = m_EditMode
    End If
End Property

Private Sub ReSetPageEditMode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ҳ��ı༭ģʽ
    '����:���˺�
    '����:2016-03-25 15:30:59
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error Resume Next
    For i = 0 To ClinicDetail.UBound
        ClinicDetail(i).EditMode = m_EditMode
    Next
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,5
Public Property Get ����Ƶ��() As Integer
    ����Ƶ�� = m_����Ƶ��
End Property

Public Property Let ����Ƶ��(ByVal New_����Ƶ�� As Integer)
    m_����Ƶ�� = New_����Ƶ��
    PropertyChanged "����Ƶ��"
    Call Set����Ƶ��(m_����Ƶ��)

End Property
Private Sub Set����Ƶ��(ByVal int����Ƶ�� As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������Ƶ��
    '���:int����Ƶ��
    '����:���˺�
    '����:2016-03-30 16:25:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Err = 0: On Error Resume Next
    For i = 0 To ClinicDetail.Count
        ClinicDetail(i).����Ƶ�� = int����Ƶ��
        mobj�����¼��(i).������Ϣ��.����Ƶ�� = int����Ƶ��
    Next
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,5
Public Property Get ҽ������() As String
    ҽ������ = m_ҽ������
End Property

Public Property Let ҽ������(ByVal New_ҽ������ As String)
    m_ҽ������ = New_ҽ������
    PropertyChanged "ҽ������"
    Call Setҽ������(m_ҽ������)

End Property
Private Sub Setҽ������(ByVal strҽ������ As String)
    '����:��������ҽ������
    '���:strҽ������
    Dim i As Integer
    Err = 0: On Error Resume Next
    For i = 0 To ClinicDetail.Count - 1
        ClinicDetail(i).ҽ������ = strҽ������
        If mobj�����¼��.Count > i Then
            mobj�����¼��(i + 1).ҽ������ = strҽ������
            mobj�����¼��(i + 1).�����������Ҽ�.ҽ������ = strҽ������
        End If
    Next
End Sub
