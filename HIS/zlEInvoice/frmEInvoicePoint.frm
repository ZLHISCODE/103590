VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEInvoicePoint 
   BorderStyle     =   0  'None
   Caption         =   "����Ʊ�ݿ�Ʊ��"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   2868
      Left            =   840
      ScaleHeight     =   2865
      ScaleWidth      =   4665
      TabIndex        =   4
      Top             =   1350
      Width           =   4668
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3225
         Left            =   2400
         MousePointer    =   9  'Size W E
         ScaleHeight     =   3225
         ScaleMode       =   0  'User
         ScaleWidth      =   22.5
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   35
      End
      Begin VB.PictureBox picFun 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   2520
         ScaleHeight     =   2655
         ScaleWidth      =   2415
         TabIndex        =   9
         Top             =   120
         Width           =   2415
         Begin XtremeSuiteControls.TabControl tbPage 
            Height          =   615
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   855
            _Version        =   589884
            _ExtentX        =   1508
            _ExtentY        =   1085
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picTree 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   2055
         TabIndex        =   7
         Top             =   120
         Width           =   2055
         Begin MSComctlLib.TreeView tvw��Ʊ�� 
            Height          =   1485
            Left            =   -360
            TabIndex        =   8
            Top             =   480
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   2619
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "ils16"
            Appearance      =   1
         End
      End
   End
   Begin VB.PictureBox pic����Ʊ�ݶ��� 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   6264
      ScaleHeight     =   1935
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
      Begin VSFlex8Ctl.VSFlexGrid vs������ϸ 
         Height          =   1080
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "��Ʊ�������ϸ"
         Top             =   240
         Width           =   1995
         _cx             =   3519
         _cy             =   1905
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEInvoicePoint.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox pic����Ʊ������ 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   6360
      ScaleHeight     =   2175
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
      Begin VB.PictureBox picSplitH 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   3000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3000
      End
      Begin VSFlex8Ctl.VSFlexGrid vs���� 
         Height          =   1080
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "��Ʊ�����"
         Top             =   1080
         Width           =   1995
         _cx             =   3519
         _cy             =   1905
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEInvoicePoint.frx":016D
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vs��Ʊ�� 
         Height          =   1080
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "��Ʊ�����"
         Top             =   0
         Width           =   1995
         _cx             =   3519
         _cy             =   1905
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEInvoicePoint.frx":02DA
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEInvoicePoint.frx":0467
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEInvoicePoint.frx":1DF9
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   3720
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
            Picture         =   "frmEInvoicePoint.frx":378B
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEInvoicePoint.frx":419D
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1032
      Left            =   0
      Top             =   648
      Width           =   528
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "������������"
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
Attribute VB_Name = "frmEInvoicePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng���볤�� As Long
Private mintColumn As Integer
Private mstrKey As String       'ǰһ�����ڵ�Ĺؼ�ֵ
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mint���� As Integer
Private mblnShowStop As Boolean  '��ʾͣ��
Private mblnShowAll As Boolean  '��ʾ�����¼�
Public mint���뷽ʽ As Integer  '0-���ͻ��˶�,1-���շ�Ա��;2-���շ�Ա+�ͻ��˶�
Private Enum mFocus
    Focus_None = 0
    Focus_��Ʊ����� = 1
    Focus_��Ʊ�� = 2
    Focus_��Ʊ����� = 3
    Focus_������ϸ = 4
End Enum
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long
Dim sngStartX As Single, sngStartY As Single    '�ƶ�ǰ����λ��

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHandler
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "ͣ��(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Add, "��������(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Modify, "�޸Ķ���(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Delete, "ɾ������(&D")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾ��ͣ�ÿ�Ʊ��(&P)", cbrControl.index)
        cbrControl.Checked = mblnShowStop
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "��ʾ�����¼�(&H)", cbrControl.index)
        cbrControl.Checked = mblnShowAll
        cbrControl.BeginGroup = True
    End With
    
    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&E)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&R)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "ͣ��(&P)", cbrControl.index + 1): cbrControl.BeginGroup = True
        .Item(cbrControl.index + 1).BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Add, "��������(&A)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Modify, "�޸Ķ���(&U)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit_Delete, "ɾ������(&D)", cbrControl.index + 1)
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add FCONTROL, Asc("P"), conMenu_Edit_Pause
        .Add FCONTROL, Asc("A"), conMenu_Edit_Audit_Add
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit_Modify
        .Add FCONTROL, Asc("D"), conMenu_Edit_Audit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID

    Case conMenu_File_Parameter '��������
        Set objfrmEInvoiceParaSet = New frmEInvoiceParaSet
        Call objfrmEInvoiceParaSet.ShowMe(Me, mlngSys, 1145)
        
         mint���뷽ʽ = zlDatabase.GetPara("��Ʊ����뷽ʽ", mlngSys, 1145, 1)
        Call Load��Ʊ�����(0)
        Call Load��Ʊ�������ϸ(0)
        
    Case conMenu_Edit_NewItem '����
        Call AddNewEInvoicePoint
    Case conMenu_Edit_Modify  '�޸�
        Call ModifyEInvoicePoint
    Case conMenu_Edit_Delete 'ɾ��
        Call DeleteEInvoicePoint
    Case conMenu_Edit_Reuse '����
        Call StartEInvoicePoint
    Case conMenu_Edit_Pause 'ͣ��
        Call StopEInvoicePoint
    Case conMenu_Edit_Audit_Add '��������
        Call Set��Ʊ�����
    Case conMenu_Edit_Audit_Modify '�޸Ķ���
        Call Set��Ʊ�����(True)
    Case conMenu_Edit_Audit_Delete 'ɾ������
        Call Delete��Ʊ�����
    Case conMenu_View_ShowStoped '��ʾͣ�õ�
         Control.Checked = Not Control.Checked
         mblnShowStop = Control.Checked
         Call load��Ʊ�����
    Case conMenu_View_ShowAll '��ʾ�����¼�
         Control.Checked = Not Control.Checked
         mblnShowAll = Control.Checked
         Call load��Ʊ�����
    Case conMenu_View_Refresh 'ˢ������
        Call load��Ʊ�����
    Case Else
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean
    Dim int���� As Integer
    Dim blnͣ�� As Boolean, bln���� As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub

    If mint���� = Focus_��Ʊ�� Then
        blnEnable = Val(vs��Ʊ��.TextMatrix(vs��Ʊ��.Row, vs��Ʊ��.ColIndex("��Ʊ��ID"))) > 0
        blnͣ�� = Val(vs��Ʊ��.TextMatrix(vs��Ʊ��.Row, vs��Ʊ��.ColIndex("ͣ��"))) = "1"
    End If
    
    If mint���� = Focus_��Ʊ����� Then
        bln���� = Val(vs����.TextMatrix(vs����.Row, vs����.ColIndex("ID"))) > 0
    End If
    If mint���� = Focus_������ϸ Then
        bln���� = Val(vs������ϸ.TextMatrix(vs������ϸ.Row, vs������ϸ.ColIndex("ID"))) > 0
    End If
    
    Select Case Control.ID
     Case conMenu_Edit_NewItem
        Control.Enabled = mint���� = Focus_��Ʊ����� Or mint���� = Focus_��Ʊ��
    Case conMenu_Edit_Modify
        If mint���� = Focus_��Ʊ�� Then
            Control.Enabled = blnEnable
        ElseIf mint���� = Focus_��Ʊ����� Then
            Control.Enabled = tvw��Ʊ��.SelectedItem.Key <> "Root"
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Delete
        If mint���� = Focus_��Ʊ�� Then
            Control.Enabled = blnEnable
        ElseIf mint���� = Focus_��Ʊ����� Then
            Control.Enabled = tvw��Ʊ��.SelectedItem.Image <> "Root"
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Reuse
        If mint���� = Focus_��Ʊ�� Then
            If Not blnEnable Then
                Control.Enabled = False
            Else
                Control.Enabled = blnͣ��
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Pause
        If mint���� = Focus_��Ʊ�� Then
            If Not blnEnable Then
                Control.Enabled = False
            Else
                Control.Enabled = Not blnͣ��
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_Edit_Audit_Add
        Control.Enabled = mint���� = Focus_��Ʊ����� Or mint���� = Focus_������ϸ
    Case conMenu_Edit_Audit_Modify, conMenu_Edit_Audit_Delete
        Control.Enabled = (mint���� = Focus_��Ʊ����� Or mint���� = Focus_������ϸ) And bln����
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel��
        Control.Enabled = False
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    Call InitPage
    mblnShowStop = GetSetting("ZLSOFT", "˽��ģ��\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ�ÿ�Ʊ��", 0)
    mblnShowAll = GetSetting("ZLSOFT", "˽��ģ��\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ�����¼���Ʊ��", 0)
    mint���뷽ʽ = zlDatabase.GetPara("��Ʊ����뷽ʽ", 100, 1145, 1)

    Call load��Ʊ�����
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Public Sub AddNewEInvoicePoint()
    '����
    Dim strKey As String
    Dim blnĩ�� As Boolean, blnRefresh As Boolean
    Dim frmEdit As New frmEInvoicePointSet

    On Error GoTo errHandle
    If tvw��Ʊ��.SelectedItem Is Nothing Then Exit Sub
    strKey = Mid(tvw��Ʊ��.SelectedItem.Key, 2)
    blnĩ�� = mint���� = Focus_��Ʊ��
    Call frmEdit.Init��Ʊ������("", strKey, blnĩ��, blnRefresh)
    If blnRefresh Then Call load��Ʊ�����
  
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteEInvoicePoint()
    'ɾ��
    On Error GoTo errHandle
    Dim strKey As String, strSQL As String
    Dim intIndex As Long
    Dim strTemp As String
    
    If mint���� = Focus_��Ʊ����� Then
        If tvw��Ʊ��.SelectedItem Is Nothing Then Exit Sub
        strKey = tvw��Ʊ��.SelectedItem.Key
        If strKey = "Root" Then Exit Sub
        strTemp = Val(Mid(tvw��Ʊ��.SelectedItem.Key, 2))
    
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "�õ���Ʊ�ݿ�Ʊ���¼�����������Ʊ�㣬����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & tvw��Ʊ��.SelectedItem.Text & "���ĵ���Ʊ�ݿ�Ʊ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        MousePointer = 11
        strSQL = "zl_����Ʊ�ݿ�Ʊ��_DELETE(" & strTemp & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        MousePointer = 0
    Else
        With vs��Ʊ��
            If .Row = 0 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID"))) = 0 Then Exit Sub
            strTemp = Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID")))
            If CheckExistDepPres(strTemp) = True Then
                MsgBox "�õ���Ʊ�ݿ�Ʊ���¼�����������Ʊ�㣬����ɾ����", vbInformation, gstrSysName
                Exit Sub
            End If
            If MsgBox("��ȷ��Ҫɾ������Ϊ��" & vs��Ʊ��.TextMatrix(vs��Ʊ��.Row, vs��Ʊ��.ColIndex("����")) & "���ĵ���Ʊ�ݿ�Ʊ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Me.MousePointer = 11
            strSQL = "Zl_����Ʊ�ݿ�Ʊ��_DELETE(" & strTemp & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Me.MousePointer = 0
        End With
    End If
    Call load��Ʊ�����
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub ModifyEInvoicePoint(Optional ByVal strID As String)
    '�޸ĵ���Ʊ�ݿ�Ʊ��
    'strID-��Ʊ��id
    Dim strKey As String
    Dim rsTemp As ADODB.Recordset
    Dim str�ϼ����� As String, blnRefresh As Boolean
    Dim strTemp As String, blnĩ�� As Boolean
    Dim frmEdit As New frmEInvoicePointSet
    
    On Error Resume Next
    blnĩ�� = mint���� = Focus_��Ʊ��
    
    If Val(strID) = 0 Then
        If mint���� = Focus_��Ʊ�� Then
            strID = vs��Ʊ��.TextMatrix(vs��Ʊ��.Row, vs��Ʊ��.ColIndex("��Ʊ��ID"))
        ElseIf mint���� = Focus_��Ʊ����� Then
            strID = Val(Mid(tvw��Ʊ��.SelectedItem.Key, 2))
            If strID = 0 Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    Call frmEdit.Init��Ʊ������(strID, , blnĩ��, blnRefresh)
    If blnRefresh Then Call load��Ʊ�����
End Sub

Public Sub StartEInvoicePoint()
    '����
    On Error GoTo errHandle
    Dim strKey As String, strSQL As String
    Dim intIndex As Long
    Dim strTemp As String
    
    With vs��Ʊ��
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID"))) = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("ͣ��")) = "" Then Exit Sub
        strTemp = Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID")))
        Me.MousePointer = 11
        strSQL = "zl_����Ʊ�ݿ�Ʊ��_Start(" & strTemp & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Me.MousePointer = 0
    End With
    
     Call load��Ʊ�����
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub StopEInvoicePoint()
    'ͣ��
    On Error GoTo errHandle
    Dim strKey As String, strSQL As String
    Dim intIndex As Long
    Dim strTemp As String

    With vs��Ʊ��
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID"))) = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("ͣ��")) = "1" Then Exit Sub
        strTemp = Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID")))
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "�õ���Ʊ�ݿ�Ʊ���¼�����������Ʊ�㣬����ͣ�á�", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("��ȷ��Ҫͣ������Ϊ��" & .TextMatrix(.Row, .ColIndex("����")) & "���ĵ���Ʊ�ݿ�Ʊ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Me.MousePointer = 11
            strSQL = "zl_����Ʊ�ݿ�Ʊ��_Stop(" & strTemp & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Me.MousePointer = 0
        End If
    End With

     Call load��Ʊ�����
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "˽��ģ��\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ�ÿ�Ʊ��", mblnShowStop
    SaveSetting "ZLSOFT", "˽��ģ��\" & mstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ�����¼���Ʊ��", mblnShowAll
     
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
End Sub

Private Sub picFun_Resize()
    On Error Resume Next
    With picFun
        tbPage.Left = 0
        tbPage.Top = 0
        tbPage.Height = .ScaleHeight
        tbPage.Width = .ScaleWidth
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    With picMain
        picTree.Left = 0
        picTree.Top = 0
        picTree.Height = .ScaleHeight
        picTree.Width = .ScaleWidth * 0.2
        picFun.Left = picTree.Width
        picFun.Top = 0
        picFun.Height = .ScaleHeight
        picFun.Width = .ScaleWidth * 0.8
        picSplit.Left = picTree.Width
        picSplit.Height = .ScaleHeight
    End With
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartX = X
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - sngStartX
        If sngTemp > 1000 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 1000 Then
            picSplit.Left = sngTemp
            picTree.Width = picSplit.Left
            picFun.Left = picSplit.Left + picSplit.Width
            picFun.Width = picMain.ScaleWidth - picFun.Left
        End If
        zlcontrol.ControlSetFocus tvw��Ʊ��
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - vs��Ʊ��.Top > 2500 And Me.ScaleHeight - (sngTemp + picSplitH.Height) > 1500 Then
            picSplitH.Top = sngTemp
            vs��Ʊ��.Height = picSplitH.Top
            vs����.Top = picSplitH.Top + picSplitH.Height
            vs����.Height = pic����Ʊ������.ScaleHeight - vs����.Top
        End If
        zlcontrol.ControlSetFocus vs��Ʊ��
    End If
End Sub

Private Sub picTree_Resize()
    On Error Resume Next
    With picTree
        tvw��Ʊ��.Left = 0
        tvw��Ʊ��.Top = 0
        tvw��Ʊ��.Height = .ScaleHeight
        tvw��Ʊ��.Width = .ScaleWidth
    End With
End Sub

Private Sub pic����Ʊ�ݶ���_Resize()
    On Error Resume Next
    With pic����Ʊ�ݶ���
        vs������ϸ.Left = 0
        vs������ϸ.Top = 0
        vs������ϸ.Width = .ScaleWidth
        vs������ϸ.Height = .ScaleHeight
    End With
End Sub

Private Sub pic����Ʊ������_Resize()
    On Error Resume Next
    With pic����Ʊ������
        vs��Ʊ��.Left = 0
        vs��Ʊ��.Top = 0
        vs��Ʊ��.Width = .ScaleWidth
        vs��Ʊ��.Height = 0.6 * .ScaleHeight
        vs����.Left = 0
        vs����.Top = vs��Ʊ��.Height
        vs����.Width = .ScaleWidth
        vs����.Height = 0.4 * .ScaleHeight
        picSplitH.Left = 0
        picSplitH.Top = vs��Ʊ��.Top + vs��Ʊ��.Height
        picSplitH.Width = .ScaleWidth
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tvw��Ʊ��.SelectedItem Is Nothing Then Exit Sub
    If Item.Caption = "Ʊ�ݶ�����ϸ" Then
        Call Load��Ʊ�������ϸ(Val(Mid(tvw��Ʊ��.SelectedItem.Key, 2)))
        picSplitH.Visible = False
    Else
        Call load��Ʊ��(tvw��Ʊ��.SelectedItem.Key)
        picSplitH.Visible = True
    End If
End Sub

Private Sub tvw��Ʊ��_DblClick()
    If tvw��Ʊ��.SelectedItem Is Nothing Then Exit Sub
    Call ModifyEInvoicePoint(Val(Mid(tvw��Ʊ��.SelectedItem.Key, 2)))
End Sub

Private Sub tvw��Ʊ��_GotFocus()
    mint���� = mFocus.Focus_��Ʊ�����
End Sub

Private Sub tvw��Ʊ��_LostFocus()
    mint���� = mFocus.Focus_None
End Sub

Private Sub tvw��Ʊ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint���� = mFocus.Focus_��Ʊ�����
    Call ShowPopup
End Sub


Public Sub tvw��Ʊ��_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errHandle
    If Node Is Nothing Then Exit Sub
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    If tbPage.Selected.Caption = "Ʊ�ݿ�Ʊ��" Then
        Call load��Ʊ��(mstrKey)
    Else
        Call Load��Ʊ�������ϸ(Val(Mid(mstrKey, 2)))
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub load��Ʊ�����()
'����:���ؿ�Ʊ�����
    Dim strSQL As String
    Dim strKey As String
    Dim rs��Ʊ�� As ADODB.Recordset
    Dim i As Integer
    Dim nod As Node
    
    mstrKey = ""
    On Error GoTo errHandle

    If Not tvw��Ʊ��.SelectedItem Is Nothing Then
        strKey = tvw��Ʊ��.SelectedItem.Key
    End If
            
    strSQL = " Select ID, �ϼ�id, ����, ����, ����, Ժ��, �ͻ���, λ��, ĩ��, ����ʱ��, ����ʱ�� From ����Ʊ�ݿ�Ʊ�� " & _
                  " Where  Nvl(ĩ��, 0) = 0 " & _
                  " Start with �ϼ�id is null connect by prior id=�ϼ�id"
    Set rs��Ʊ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    tvw��Ʊ��.Nodes.Clear
    tvw��Ʊ��.Nodes.Add , , "Root", "���п�Ʊ��", "Root", "Root"
    tvw��Ʊ��.Nodes("Root").Sorted = True
    
    Do Until rs��Ʊ��.EOF
            
        If IsNull(rs��Ʊ��("�ϼ�id")) Then
            tvw��Ʊ��.Nodes.Add "Root", tvwChild, "_" & rs��Ʊ��("id"), "��" & rs��Ʊ��("����") & "��" & rs��Ʊ��("����"), "Dept", "Dept"
        Else
            tvw��Ʊ��.Nodes.Add "_" & rs��Ʊ��("�ϼ�id"), tvwChild, "_" & rs��Ʊ��("id"), "��" & rs��Ʊ��("����") & "��" & rs��Ʊ��("����"), "Dept", "Dept"
        End If
        tvw��Ʊ��.Nodes("_" & rs��Ʊ��("id")).Sorted = True
        rs��Ʊ��.MoveNext
    Loop

    On Error Resume Next
    Set nod = tvw��Ʊ��.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvw��Ʊ��.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
        tvw��Ʊ��_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvw��Ʊ��_NodeClick nod
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub load��Ʊ��(ByVal str�ϼ�ID As String)
'����:���ؿ�Ʊ��
'����:str�ϼ�ID :����Ʊ�ݿ�Ʊ��.�ϼ�id
    Dim strSQL As String
    Dim rs��Ʊ�� As ADODB.Recordset
    Dim blnͣ�� As Boolean
    Dim lng��Ʊ��id As Long, i As Integer
    Dim strͣ�� As String, intRow As Integer
    
    On Error GoTo errHandle
    Call Load��Ʊ�����(0)
    
    If Not vs��Ʊ��.Row = 0 Then
        '����ԭ�м�ֵ
        lng��Ʊ��id = Val(vs��Ʊ��.TextMatrix(vs��Ʊ��.Row, 0))
    End If
    
    If Not mblnShowStop Then
        strͣ�� = " And (A.����ʱ�� is null or A.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    If mblnShowAll Then
        strSQL = "Select a.*, b.���� As ����" & vbNewLine & _
                      "From (Select a.Id, a.�ϼ�id, b.���� As �ϼ�, a.����, a.����, a.����, a.λ��, a.�ͻ���, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��," & vbNewLine & _
                      "              To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��, a.����id, a.Ժ��" & vbNewLine & _
                      "       From ����Ʊ�ݿ�Ʊ�� A, ����Ʊ�ݿ�Ʊ�� B" & vbNewLine & _
                      "       Where a.�ϼ�id = b.Id(+) And Nvl(a.ĩ��, 0) = 1" & strͣ�� & vbNewLine & _
                      "       Connect By Prior a.Id = a.�ϼ�id start with " & IIf(str�ϼ�ID = "Root", "A.�ϼ�ID is null ", "A.�ϼ�ID = [1]") & _
                      "       ) A, ���ű� B" & vbNewLine & _
                      "Where a.����id = b.Id(+)"
    Else
         strSQL = "Select a.Id, a.�ϼ�id, c.���� As �ϼ�, a.����, a.����, a.����, a.λ��, a.�ͻ���, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��," & vbNewLine & _
                      "              To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��, a.����id, a.Ժ��,b.���� As ����" & vbNewLine & _
                      "       From ����Ʊ�ݿ�Ʊ�� A, ���ű� B,����Ʊ�ݿ�Ʊ�� C" & vbNewLine & _
                      "       Where Nvl(a.ĩ��, 0) = 1 And a.����id = b.Id(+) And a.�ϼ�id = c.Id(+) " & strͣ�� & IIf(str�ϼ�ID = "Root", " And A.�ϼ�ID is null ", " And A.�ϼ�ID = [1]")
    End If
    Set rs��Ʊ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(str�ϼ�ID, 2)))
    vs��Ʊ��.Clear 1: vs��Ʊ��.Rows = 2
    If rs��Ʊ��.EOF Then Exit Sub
    With vs��Ʊ��
        .Rows = rs��Ʊ��.RecordCount + 1
        For i = 1 To rs��Ʊ��.RecordCount
            .TextMatrix(i, .ColIndex("��Ʊ��ID")) = rs��Ʊ��!ID
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs��Ʊ��!����)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs��Ʊ��!����)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs��Ʊ��!����)
            .TextMatrix(i, .ColIndex("�ͻ���")) = Nvl(rs��Ʊ��!�ͻ���)
            .TextMatrix(i, .ColIndex("λ��")) = Nvl(rs��Ʊ��!λ��)
            .TextMatrix(i, .ColIndex("����ʱ��")) = Nvl(rs��Ʊ��!����ʱ��)
            .TextMatrix(i, .ColIndex("����ʱ��")) = Nvl(rs��Ʊ��!����ʱ��)
            .TextMatrix(i, .ColIndex("�ϼ�")) = Nvl(rs��Ʊ��!�ϼ�)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs��Ʊ��!����)
            .TextMatrix(i, .ColIndex("Ժ��")) = Nvl(rs��Ʊ��!Ժ��)
             If Not CDate(IIf(IsNull(rs��Ʊ��("����ʱ��")), CDate("3000/1/1"), rs��Ʊ��("����ʱ��"))) = CDate("3000/1/1") Then
                .Cell(flexcpForeColor, i, .ColIndex("����"), i, .ColIndex("Ժ��")) = RGB(255, 0, 0)
                .TextMatrix(i, .ColIndex("ͣ��")) = "1"
            End If
            rs��Ʊ��.MoveNext
        Next
        intRow = .FindRow(lng��Ʊ��id, 0, .ColIndex("��Ʊ��ID"), , True)
        If intRow > 0 Then .Row = intRow
        If Val(.TextMatrix(1, .ColIndex("��Ʊ��ID"))) > 0 Then
            Call Load��Ʊ�����(Val(.TextMatrix(.Row, .ColIndex("��Ʊ��ID"))))
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckExistDepPres(ByVal lng�ϼ�id As Long) As Boolean
    '���õ���Ʊ�ݿ�Ʊ�����Ƿ����������Ʊ��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 1 From ����Ʊ�ݿ�Ʊ�� " & _
        " Where �ϼ�id =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ʊ�ݿ�Ʊ��", lng�ϼ�id)
    
    If rsTemp.RecordCount > 0 Then
        CheckExistDepPres = True
        Exit Function
    End If
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Public Sub Load��Ʊ�����(ByVal lng��Ʊ��id As Long)
    '���ݿ�Ʊ��id���ؿ�Ʊ�������Ϣ
    Dim strSQL As String, i As Integer
    Dim rs��Ʊ����� As New ADODB.Recordset
    
    With vs����
        .ColHidden(.ColIndex("�շ�Ա")) = mint���뷽ʽ = 0
        .ColHidden(.ColIndex("�շ�Ա���")) = mint���뷽ʽ = 0
        .ColHidden(.ColIndex("�շ�Ա��������")) = mint���뷽ʽ = 0
        .ColHidden(.ColIndex("�ͻ���")) = mint���뷽ʽ = 1
        .ColHidden(.ColIndex("����")) = mint���뷽ʽ = 1
        .ColHidden(.ColIndex("��;")) = mint���뷽ʽ = 1
    End With
    vs����.Clear 1: vs����.Rows = 2
    If lng��Ʊ��id = 0 Then Exit Sub
    strSQL = "Select a.Id As ��Ʊ��id, b.id,a.���� As ��Ʊ��, b.��Աid, c.���� As �շ�Ա, c.��� as �շ�Ա���,g.���� As �շ�Ա��������, b.�ͻ���, e.����, e.��;" & vbNewLine & _
                    "From ����Ʊ�ݿ�Ʊ�� A, Ʊ�ݿ�Ʊ����� B, ��Ա�� C, ������Ա D, zlClients E, ���ű� G" & vbNewLine & _
                    "Where a.Id = b.��Ʊ��id(+) And b.��Աid = c.Id(+) And b.�ͻ��� = e.����վ(+) And c.Id = d.��Աid(+) And d.ȱʡ(+) = 1 And" & vbNewLine & _
                    "      d.����id = g.Id(+) And a.Id = [1] "
    Set rs��Ʊ����� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ʊ��id)
    If rs��Ʊ�����.EOF Then Exit Sub
    With vs����
        .Rows = rs��Ʊ�����.RecordCount + 1
        For i = 1 To rs��Ʊ�����.RecordCount
            .TextMatrix(i, .ColIndex("��Ʊ��ID")) = rs��Ʊ�����!��Ʊ��id
            .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rs��Ʊ�����!ID))
            .TextMatrix(i, .ColIndex("��Ʊ��")) = Nvl(rs��Ʊ�����!��Ʊ��)
            .TextMatrix(i, .ColIndex("��Աid")) = Nvl(rs��Ʊ�����!��Աid)
            .TextMatrix(i, .ColIndex("�շ�Ա")) = Nvl(rs��Ʊ�����!�շ�Ա)
            .TextMatrix(i, .ColIndex("�շ�Ա���")) = Nvl(rs��Ʊ�����!�շ�Ա���)
            .TextMatrix(i, .ColIndex("�շ�Ա��������")) = Nvl(rs��Ʊ�����!�շ�Ա��������)
            .TextMatrix(i, .ColIndex("�ͻ���")) = Nvl(rs��Ʊ�����!�ͻ���)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs��Ʊ�����!����)
            .TextMatrix(i, .ColIndex("��;")) = Nvl(rs��Ʊ�����!��;)
            rs��Ʊ�����.MoveNext
        Next
    End With
End Sub

Public Sub Load��Ʊ�������ϸ(ByVal lng�ϼ�id As Long)
    '���ݿ�Ʊ��id���ؿ�Ʊ�������Ϣ
    Dim strSQL As String, i As Integer
    Dim rs��Ʊ����� As New ADODB.Recordset
    
    vs������ϸ.Clear 1: vs������ϸ.Rows = 2
    With vs������ϸ
        .ColHidden(.ColIndex("�շ�Ա")) = mint���뷽ʽ = 0
        .ColHidden(.ColIndex("�շ�Ա���")) = mint���뷽ʽ = 0
        .ColHidden(.ColIndex("�շ�Ա��������")) = mint���뷽ʽ = 0
        .ColHidden(.ColIndex("�ͻ���")) = mint���뷽ʽ = 1
        .ColHidden(.ColIndex("����")) = mint���뷽ʽ = 1
        .ColHidden(.ColIndex("��;")) = mint���뷽ʽ = 1
    End With
    If mblnShowAll Then
        strSQL = "Select  a.��Ʊ��id, a.��Ʊ��, f.Id, f.��Աid, f.�ͻ���, b.���� As �շ�Ա, b.��� As �շ�Ա���, e.���� As �շ�Ա��������, d.����, d.��; " & vbNewLine & _
                        "From(Select a.Id As ��Ʊ��id, a.���� As ��Ʊ�� From ����Ʊ�ݿ�Ʊ�� A Where a.ĩ�� = 1 Connect By Prior a.Id = a.�ϼ�id" & vbNewLine & _
                        "Start With " & IIf(Val(lng�ϼ�id) = 0, "A.�ϼ�ID is null )", "A.�ϼ�ID = [1])") & "A, ��Ա�� B, ������Ա C, zlClients D, ���ű� E, Ʊ�ݿ�Ʊ����� F " & vbNewLine & _
                        "Where f.��Աid = b.Id(+) And f.�ͻ��� = d.����վ(+) And f.Id = c.��Աid(+) And c.ȱʡ(+) = 1  And " & vbNewLine & _
                        "      c.����id = e.Id(+) And a.��Ʊ��id = f.��Ʊ��id(+)"
    Else
        strSQL = "Select a.��Ʊ��id, a.��Ʊ��, f.Id, f.��Աid, f.�ͻ���, b.���� As �շ�Ա, b.��� As �շ�Ա���, e.���� As �շ�Ա��������, d.����, d.��; " & vbNewLine & _
                        "From(Select a.Id As ��Ʊ��id, a.���� As ��Ʊ�� From ����Ʊ�ݿ�Ʊ�� A Where  a.ĩ�� = 1 " & vbNewLine & _
                        "And " & IIf(Val(lng�ϼ�id) = 0, "A.�ϼ�ID is null )", "A.�ϼ�ID = [1])") & "A, ��Ա�� B, ������Ա C, zlClients D, ���ű� E, Ʊ�ݿ�Ʊ����� F" & vbNewLine & _
                        "Where f.��Աid = b.Id(+) And f.�ͻ��� = d.����վ(+) And f.Id = c.��Աid(+) And c.ȱʡ(+) = 1  And  " & vbNewLine & _
                        "      c.����id = e.Id(+) And a.��Ʊ��id = f.��Ʊ��id(+)"
    End If
    Set rs��Ʊ����� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ϼ�id)
    If rs��Ʊ�����.EOF Then Exit Sub
    With vs������ϸ
        .Rows = rs��Ʊ�����.RecordCount + 1
        For i = 1 To rs��Ʊ�����.RecordCount
            .TextMatrix(i, .ColIndex("��Ʊ��ID")) = rs��Ʊ�����!��Ʊ��id
            .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(rs��Ʊ�����!ID))
            .TextMatrix(i, .ColIndex("��Ʊ��")) = Nvl(rs��Ʊ�����!��Ʊ��)
            .TextMatrix(i, .ColIndex("��Աid")) = Nvl(rs��Ʊ�����!��Աid)
            .TextMatrix(i, .ColIndex("�շ�Ա")) = Nvl(rs��Ʊ�����!�շ�Ա)
            .TextMatrix(i, .ColIndex("�շ�Ա���")) = Nvl(rs��Ʊ�����!�շ�Ա���)
            .TextMatrix(i, .ColIndex("�շ�Ա��������")) = Nvl(rs��Ʊ�����!�շ�Ա��������)
            .TextMatrix(i, .ColIndex("�ͻ���")) = Nvl(rs��Ʊ�����!�ͻ���)
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs��Ʊ�����!����)
            .TextMatrix(i, .ColIndex("��;")) = Nvl(rs��Ʊ�����!��;)
            rs��Ʊ�����.MoveNext
        Next
    End With
End Sub

Private Sub Set��Ʊ�����(Optional ByVal blnModify As Boolean)
    '���ݿ�Ʊ��id���ÿ�Ʊ�������Ϣ
    Dim lng��Ʊ��id As Long, lngID As Long
    Dim frmEdit As New frmEInvoicePointSet
    Dim bln���� As Boolean, blnRefresh As Boolean
    
    If mint���� <> Focus_��Ʊ����� And mint���� <> Focus_������ϸ Then Exit Sub
    bln���� = mint���� = Focus_��Ʊ�����
    If bln���� Then
        lng��Ʊ��id = Val(vs����.TextMatrix(vs����.Row, vs����.ColIndex("��Ʊ��id")))
        lngID = Val(vs����.TextMatrix(vs����.Row, vs����.ColIndex("id")))
    Else
        lng��Ʊ��id = Val(vs������ϸ.TextMatrix(vs������ϸ.Row, vs������ϸ.ColIndex("��Ʊ��id")))
        lngID = Val(vs������ϸ.TextMatrix(vs������ϸ.Row, vs������ϸ.ColIndex("ID")))
    End If
    If lng��Ʊ��id = 0 Then Exit Sub
    If Not blnModify Then
        lngID = 0
    Else
        If lngID = 0 Then Exit Sub
    End If
    Call frmEdit.Init��Ʊ�����(mint���뷽ʽ, lng��Ʊ��id, lngID, blnRefresh)
    If Not blnRefresh Then Exit Sub
    If bln���� Then
        Call Load��Ʊ�����(lng��Ʊ��id)
    Else
        Call Load��Ʊ�������ϸ(Val(Mid(mstrKey, 2)))
    End If
End Sub

Private Sub Delete��Ʊ�����()
    'ɾ����Ʊ�������Ϣ
    Dim lngID As Long, lng��Ʊ��id As Long
    Dim strSQL As String, bln���� As Boolean
    
    If mint���� <> Focus_��Ʊ����� And mint���� <> Focus_������ϸ Then Exit Sub
    bln���� = mint���� = Focus_��Ʊ�����
    If bln���� Then
        lngID = Val(vs����.TextMatrix(vs����.Row, vs����.ColIndex("id")))
        lng��Ʊ��id = Val(vs����.TextMatrix(vs����.Row, vs����.ColIndex("��Ʊ��id")))
    Else
        lngID = Val(vs������ϸ.TextMatrix(vs������ϸ.Row, vs������ϸ.ColIndex("id")))
        lng��Ʊ��id = Val(vs������ϸ.TextMatrix(vs������ϸ.Row, vs������ϸ.ColIndex("��Ʊ��id")))
    End If
    If lng��Ʊ��id = 0 Then Exit Sub
    If lngID = 0 Then Exit Sub
    strSQL = "Zl_Ʊ�ݿ�Ʊ�����_Update(2," & lngID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ�ݿ�Ʊ�����")
    If bln���� Then
        Call Load��Ʊ�����(lng��Ʊ��id)
    Else
        Call Load��Ʊ�������ϸ(Val(Mid(mstrKey, 2)))
    End If
End Sub

Private Sub vs������ϸ_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs������ϸ, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs������ϸ_DblClick()
    '����
    Dim blnModify As Boolean
    If vs������ϸ.Row = 0 Then Exit Sub
    blnModify = Val(vs������ϸ.TextMatrix(vs������ϸ.Row, vs������ϸ.ColIndex("id"))) > 0
    Call Set��Ʊ�����(blnModify)
End Sub

Private Sub vs������ϸ_GotFocus()
    mint���� = Focus_������ϸ
    If vs������ϸ.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs������ϸ, &HFFEBD7
End Sub

Private Sub vs������ϸ_LostFocus()
    mint���� = Focus_None
    If vs������ϸ.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs������ϸ
    OS.OpenIme False
End Sub

Private Sub vs������ϸ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint���� = mFocus.Focus_������ϸ
    Call ShowPopup
End Sub

Private Sub vs����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs����, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs����_DblClick()
    '����
    Dim blnModify As Boolean
    If vs����.Row = 0 Then Exit Sub
    blnModify = Val(vs����.TextMatrix(1, vs����.ColIndex("id"))) > 0
    Call Set��Ʊ�����(blnModify)
End Sub

Private Sub vs����_GotFocus()
    mint���� = mFocus.Focus_��Ʊ�����
    If vs����.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs����, &HFFEBD7
End Sub

Private Sub vs����_LostFocus()
    mint���� = mFocus.Focus_None
    If vs����.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs����
    OS.OpenIme False
End Sub

Private Sub vs����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint���� = mFocus.Focus_��Ʊ�����
    Call ShowPopup
End Sub

Private Sub ShowPopup()
    '��ʾ�����˵�
    Dim objPopup As CommandBarPopup
    Err = 0: On Error GoTo ErrHandler
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    Dim objItem As TabControlItem
    With tbPage
        Set objItem = .InsertItem(1, "Ʊ�ݿ�Ʊ��", pic����Ʊ������.hWnd, 0)
        objItem.Tag = 1
        Set objItem = .InsertItem(2, "Ʊ�ݶ�����ϸ", pic����Ʊ�ݶ���.hWnd, 0)
        objItem.Tag = 2
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    tbPage.Item(0).Selected = True
End Sub

Private Sub vs��Ʊ��_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not (NewRow = 0 Or OldRow = 0) Then zl_VsGridRowChange vs��Ʊ��, OldRow, NewRow, OldCol, NewCol
    With vs��Ʊ��
        If NewRow = 0 Then Exit Sub
        If Val(.TextMatrix(NewRow, .ColIndex("��Ʊ��id"))) = 0 Then Exit Sub
        If .TextMatrix(NewRow, .ColIndex("ͣ��")) = "1" Then
            .Cell(flexcpForeColor, NewRow, .ColIndex("����"), NewRow, .ColIndex("Ժ��")) = RGB(255, 0, 0)
        Else
            .Cell(flexcpForeColor, NewRow, .ColIndex("����"), NewRow, .ColIndex("Ժ��")) = &H80000008
        End If
        Call Load��Ʊ�����(Val(.TextMatrix(NewRow, .ColIndex("��Ʊ��id"))))
    End With
End Sub

Private Sub vs��Ʊ��_DblClick()
    With vs��Ʊ��
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("��Ʊ��id"))) = 0 Then Exit Sub
        Call ModifyEInvoicePoint(Val(.TextMatrix(.Row, .ColIndex("��Ʊ��id"))))
    End With
End Sub

Private Sub vs��Ʊ��_GotFocus()
    mint���� = mFocus.Focus_��Ʊ��
    If vs��Ʊ��.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs��Ʊ��, &HFFEBD7
End Sub

Private Sub vs��Ʊ��_LostFocus()
    mint���� = mFocus.Focus_None
    If vs��Ʊ��.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs��Ʊ��
    OS.OpenIme False
End Sub

Private Sub vs��Ʊ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    mint���� = mFocus.Focus_��Ʊ��
    Call ShowPopup
End Sub


