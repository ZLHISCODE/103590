VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmBlackListMainFun 
   BorderStyle     =   0  'None
   Caption         =   "���˲�����¼"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFunBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   3300
      ScaleHeight     =   4965
      ScaleWidth      =   3555
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Width           =   3555
      Begin XtremeSuiteControls.ShortcutBar stbFunc 
         Height          =   4155
         Left            =   60
         TabIndex        =   7
         Top             =   210
         Width           =   3225
         _Version        =   589884
         _ExtentX        =   5689
         _ExtentY        =   7329
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBaseSetBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4905
      Left            =   45
      ScaleHeight     =   4905
      ScaleWidth      =   2985
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   465
      Width           =   2985
      Begin XtremeSuiteControls.TaskPanel tplFunBase 
         Height          =   3045
         Left            =   540
         TabIndex        =   4
         Top             =   1110
         Width           =   1815
         _Version        =   589884
         _ExtentX        =   3201
         _ExtentY        =   5371
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
      Begin XtremeSuiteControls.ShortcutCaption sccFunBase 
         Height          =   360
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   2505
         _Version        =   589884
         _ExtentX        =   4419
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picRecordBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   7380
      ScaleHeight     =   4140
      ScaleWidth      =   3345
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   585
      Width           =   3345
      Begin MSComctlLib.TreeView tvwType 
         Height          =   1425
         Left            =   75
         TabIndex        =   1
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2514
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgPlan16"
         Appearance      =   0
         OLEDragMode     =   1
      End
      Begin XtremeSuiteControls.ShortcutCaption stcRecord 
         Height          =   360
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   3165
         _Version        =   589884
         _ExtentX        =   5583
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "������Ϊ����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList imgPlan16 
      Left            =   9435
      Top             =   5850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackListMainFun.frx":0000
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackListMainFun.frx":059A
            Key             =   "Type"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgIcons 
      Bindings        =   "frmBlackListMainFun.frx":0B34
      Left            =   135
      Top             =   5760
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBlackListMainFun.frx":0B48
   End
End
Attribute VB_Name = "frmBlackListMainFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Form
Private mcbsMain As Object            'CommandBar�ؼ�
Private mstrPrivs As String
Private mlngModule As Long

Private Const M_BASE_ID = 10    '������
Private Const M_RECORD_ID = 20  '������¼
Private mblnNotClick As Boolean
Public Event zlActivate(ByVal frmSubForm As Form) '�¼�����
Public Event SelectedChange(ByVal bytFunMode As gEM_BlackListFun, ByVal strBlackLitType As String) '����ѡ��ı�ʱ



Public Sub zlInitComm(frmMain As Form, cbsThis As Object, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿ�
    '���:objPati-����������
    '     cbsThis-�˵�����
    '     strPrivs-Ȩ�޴�
    '     lngModule-ģ���
    '����:���˺�
    '����:2018-11-08 11:28:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    On Error GoTo errHandle
    Set mfrmMain = frmMain: Set mcbsMain = cbsThis
    mstrPrivs = strPrivs: mlngModule = lngModule
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub Form_Activate()
    RaiseEvent zlActivate(Me)
End Sub

Private Sub InitTaskPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ŀ������
    '����:���˺�
    '����:2018-11-08 11:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tplGroup As TaskPanelGroup
    On Error GoTo errHandle
    With tplFunBase
          .Behaviour = xtpTaskPanelBehaviourList
          .HotTrackStyle = xtpTaskPanelHighlightItem
          .SelectItemOnFocus = True
          .Icons.AddIcons imgIcons.Icons
          .SetIconSize 32, 32
          .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
          .SetMargins 1, 0, 0, 1, 2
          Set tplGroup = .Groups.Add(10, "��������")
          tplGroup.Items.Add Em_Pane_Type, "������Ϊ����", xtpTaskItemTypeLink, 12
          tplGroup.Items.Add Em_Pane_Reason, "������Ϊ����ԭ��", xtpTaskItemTypeLink, 13
          tplGroup.CaptionVisible = False
          tplGroup.Expanded = True
      End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitShortcutBar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������
    '����:���˺�
    '����:2018-11-08 11:45:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    With stbFunc
        .Icons = imgIcons.Icons
        'ͼ����������ID��ͬ
        .AddItem M_RECORD_ID, "������¼", picRecordBack.hwnd
        .AddItem M_BASE_ID, "��������", picBaseSetBack.hwnd
        .ExpandedLinesCount = .ItemCount 'Ĭ��չ��
        .Tag = M_BASE_ID
        mblnNotClick = True
        .Selected = .Item(1) 'Ҫ�л�һ�£���֤�ؼ��󶨵�λ\
        mblnNotClick = False
        .Selected = .Item(0)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2018-11-08 11:43:21
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    Call LoadTypeToTreeCtrl '������Ϊ�����ؼ�
    Call InitTaskPanel '��ʼ��������Ŀ������
    Call InitShortcutBar '��ʼ��������
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    Call InitFace   '��ʼ������ؼ�
End Sub
Public Sub zlRefresh(Optional blnNotClick As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ�¹��ܲ˵�
    '����:���˺�
    '����:2018-11-09 16:06:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call LoadTypeToTreeCtrl(blnNotClick)
    
End Sub
Private Function LoadTypeToTreeCtrl(Optional blnNotClick As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ݸ����Ϳؼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 14:14:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objNode As Node
    
    Err = 0: On Error GoTo errHandle
    
    tvwType.Nodes.Clear
    Set objNode = tvwType.Nodes.Add(, , "Root", "������Ϊ����", "Root")
    objNode.Expanded = True
    objNode.Selected = True
    
    strSQL = "Select ����,����,���� From ������Ϊ���� Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Set objNode = tvwType.Nodes.Add("Root", 4, "K" & Nvl(!����), Nvl(!����), "Type", "Type")
            objNode.Expanded = True
            objNode.Tag = Nvl(!����)
            .MoveNext
        Loop
    End With
    If Not blnNotClick Then Call tvwType_NodeClick(tvwType.SelectedItem)
    LoadTypeToTreeCtrl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    picFunBack.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function RefreshBlackListData(Optional ByVal strBlackListType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ�²�����¼����
    '���:strBlackListType-������Ϊ���,Ϊ�ձ�ʾ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 14:35:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objNode As Node
    Dim strDeletedPreviouNodeKey As String
    
    Err = 0: On Error GoTo errHandle
    If stbFunc.Selected.ID <> M_RECORD_ID Then Exit Function
    Call LoadTypeToTreeCtrl '���¼��س����
    Call tvwType_GotFocus
    RefreshBlackListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Set mcbsMain = Nothing
    Set mfrmMain = Nothing
End Sub

Private Sub picBaseSetBack_Resize()
    Err = 0: On Error Resume Next
    sccFunBase.Move 0, -10, picBaseSetBack.ScaleWidth
    With tplFunBase
        .Left = 0
        .Top = sccFunBase.Top + sccFunBase.Height
        .Width = picBaseSetBack.ScaleWidth - .Left
        .Height = picBaseSetBack.ScaleHeight - .Top
    End With
End Sub

Private Sub picRecordBack_Resize()
    Dim sngTop As Single
    
    Err = 0: On Error Resume Next
    stcRecord.Move 0, -10, picRecordBack.ScaleWidth
    
    sngTop = stcRecord.Top + stcRecord.Height
    tvwType.Move 0, sngTop, picRecordBack.ScaleWidth, picRecordBack.Height - sngTop
End Sub


Private Sub picFunBack_Resize()
    Err = 0: On Error Resume Next
    stbFunc.Move 0, 0, picFunBack.ScaleWidth, picFunBack.ScaleHeight
End Sub

Private Sub stbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub stbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    Dim tplGroup As TaskPanelGroup, tplItem As TaskPanelGroupItem, tplItemWork As TaskPanelGroupItem
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandle
    
    If mblnNotClick Then Exit Sub
    
    If Val(stbFunc.Tag) = Item.ID Then Exit Sub
    
    tvwType.Tag = ""
    
    '����Ĭ��ѡ�нڵ�
    picBaseSetBack.Visible = False
    picRecordBack.Visible = False
    
    If Item.ID = M_BASE_ID Then
        stbFunc.Tag = M_BASE_ID
        If tplFunBase.Tag = "" Then tplFunBase.Tag = "������Ϊ����"
        For Each tplGroup In tplFunBase.Groups
            For Each tplItem In tplGroup.Items
                If tplItem.Caption = tplFunBase.Tag Then Set tplItemWork = tplItem
                If tplFunBase.Tag = tplItem.Caption Then
                    tplItem.Selected = True: blnFind = True
                    tplFunBase.Tag = "": tplFunBase_ItemClick tplItem
                Else
                    tplItem.Selected = False
                End If
            Next
        Next
        
        If blnFind = False Then
            'ȱʡѡ���ϰ�ʱ��
            tplItemWork.Selected = True
            tplFunBase.Tag = "": tplFunBase_ItemClick tplItemWork
        End If
        picBaseSetBack.Visible = True
    Else
        stbFunc.Tag = M_RECORD_ID
        picRecordBack.Visible = True
        Call tvwType_GotFocus
        If tvwType.SelectedItem Is Nothing And tvwType.Nodes.Count <> 0 Then
            tvwType.Nodes("Root").Selected = True
        End If
        If Not tvwType.SelectedItem Is Nothing Then
         tvwType_NodeClick tvwType.SelectedItem
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ExcuteBlackListFun(ByVal strFunCaption As String, Optional strBlackListType As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�в�����¼��ع���
    '���:strFunCaption-��������
    '     strBlackListType-��Ϊ��������
    '����:���˺�
    '����:2018-11-08 14:56:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Select Case strFunCaption
    Case "������Ϊ����"
        RaiseEvent SelectedChange(Em_Pane_Type, "")
    Case "������Ϊ����ԭ��"
        RaiseEvent SelectedChange(Em_Pane_Reason, "")
    Case "������¼����"
        RaiseEvent SelectedChange(Em_Pane_Record, strBlackListType)
        If tvwType.Enabled And tvwType.Visible Then tvwType.SetFocus
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub tplFunBase_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Err = 0: On Error GoTo errHandle
    
    If tplFunBase.Tag = Item.Caption Then Exit Sub
    tplFunBase.Tag = Item.Caption
    Call ExcuteBlackListFun(Item.Caption)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub tvwType_GotFocus()
    '
    
End Sub

Private Sub tvwType_NodeClick(ByVal Node As MSComctlLib.Node)
    Err = 0: On Error GoTo errHandle
    If tvwType.Tag = Node.Key Then Exit Sub
    
    tvwType.Tag = Node.Key
    tvwType.HideSelection = False
    Call ExcuteBlackListFun("������¼����", Node.Tag)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
   

