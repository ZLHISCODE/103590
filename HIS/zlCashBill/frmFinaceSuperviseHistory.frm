VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFinaceSuperviseHistory 
   BorderStyle     =   0  'None
   Caption         =   "��ʷ�տ���Ϣ"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRollingCurtain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   5715
      ScaleHeight     =   3090
      ScaleWidth      =   3825
      TabIndex        =   13
      Top             =   4575
      Width           =   3825
      Begin VB.PictureBox picImgPlanRC 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   17
         Top             =   45
         Width           =   210
         Begin VB.Image imgColPlanRC 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseHistory.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   1800
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseHistory.frx":054E
         ScrollTrack     =   -1  'True
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
         OutlineBar      =   1
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   225
      ScaleHeight     =   2520
      ScaleWidth      =   3435
      TabIndex        =   11
      Top             =   4125
      Width           =   3435
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   -15
         TabIndex        =   12
         Top             =   -15
         Width           =   2865
         _Version        =   589884
         _ExtentX        =   5054
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picCollect 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   120
      ScaleHeight     =   2490
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   510
      Width           =   10170
      Begin VB.PictureBox picImgPlanCollect 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   16
         Top             =   1035
         Width           =   210
         Begin VB.Image imgColPlanCollect 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinaceSuperviseHistory.frx":061B
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   15
            Width           =   195
         End
      End
      Begin VB.TextBox txtNO 
         Height          =   345
         Left            =   1020
         TabIndex        =   10
         Top             =   540
         Width           =   3360
      End
      Begin VB.ComboBox cboDate 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "���¹�������(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7995
         TabIndex        =   1
         Top             =   95
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCollect 
         Height          =   1800
         Left            =   0
         TabIndex        =   2
         Top             =   900
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinaceSuperviseHistory.frx":0B69
         ScrollTrack     =   -1  'True
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
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   2325
         TabIndex        =   3
         Top             =   135
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   42532867
         CurrentDate     =   41520
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   4785
         TabIndex        =   4
         Top             =   135
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   42532867
         CurrentDate     =   41520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   5625
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         Caption         =   "�տ��"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   607
         Width           =   840
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4485
         TabIndex        =   7
         Top             =   187
         Width           =   225
      End
      Begin VB.Label lblHistoryDate 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   187
         Width           =   840
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2355
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFinaceSuperviseHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private Enum mPaneIndex
    EM_PN_CollectList = 270101  '�տ��б�
    EM_PN_DetailList = 270102   '�տƱ�ݻ���
End Enum
Private mblnNotBrush As Boolean '��ˢ������
Private mobjChargeBill As clsChargeBill
Private mlngCollectID As Long '�տ��տ�ID
Private mstrCollectNO As String   '�տ�ݺ�
Private mblnDel As Boolean   '�Ƿ�������
Private mint��¼���� As Integer
Private Enum mPgIndex
    EM_PG_�տƱ�� = 250101
    EM_PG_������Ϣ = 250102
End Enum
Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:���˺�
    '����:2013-09-09 14:41:46
    '˵��:���ش����,��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
End Sub
Private Sub InitPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2013-09-22 17:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(EM_PG_�տƱ��, "�տƱ����Ϣ", mobjChargeBill.GetChargeAndBillTotalForm.hWnd, 0)
    objItem.Tag = EM_PG_�տƱ��
    Set objItem = tbPage.InsertItem(EM_PG_������Ϣ, "������Ϣ", picRollingCurtain.hWnd, 0)
    objItem.Tag = EM_PG_������Ϣ
    With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-09-22 17:48:09
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngHeight  As Long
    With dkpMan
        lngHeight = 2490 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(mPaneIndex.EM_PN_CollectList, 400, lngHeight, DockRightOf, Nothing)
        objPane.Title = "�տ���Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picCollect.hWnd
        objPane.MinTrackSize.Height = lngHeight * 0.9
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_DetailList, 400, 400, DockBottomOf, objPane)
        objPane.Title = "": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:���˺�
    '����:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsCollect
        .Clear
        .Rows = 2: .Cols = 15: i = 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "��¼����": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "�տ��": i = i + 1
        .TextMatrix(0, i) = "�շ�Ա": i = i + 1
        .TextMatrix(0, i) = "�տ��": i = i + 1
        .TextMatrix(0, i) = "�տ�˵��": i = i + 1
        .TextMatrix(0, i) = "��Ԥ����": i = i + 1
        .TextMatrix(0, i) = "����ϼ�": i = i + 1
        .TextMatrix(0, i) = "����ϼ�": i = i + 1
        .TextMatrix(0, i) = "�����տ���": i = i + 1
        .TextMatrix(0, i) = "�����տ�ʱ��": i = i + 1
        .TextMatrix(0, i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .ColData(0) = "-1|1"
       For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "�տ��" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "��¼����" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "�տ��" Or .ColKey(i) = "������" Or .ColKey(i) = "����ʱ��" Then .ColData(i) = "1|0"
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "�տ��" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsCollect, Me.Name, "�տ���Ϣ�б�", False
    End With
    With vsRollingCurtain
        .Clear
        .Rows = 2: .Cols = 19: i = 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "���ʵ���": i = i + 1
        .TextMatrix(0, i) = "�������": i = i + 1
        .TextMatrix(0, i) = "��ʼʱ��": i = i + 1
        .TextMatrix(0, i) = "��ֹʱ��": i = i + 1
        .TextMatrix(0, i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .TextMatrix(0, i) = "�տ��": i = i + 1
        .TextMatrix(0, i) = "����˵��": i = i + 1
        .TextMatrix(0, i) = "��Ԥ����": i = i + 1
        .TextMatrix(0, i) = "����ϼ�": i = i + 1
        .TextMatrix(0, i) = "����ϼ�": i = i + 1
        .TextMatrix(0, i) = "С���տ���": i = i + 1
        .TextMatrix(0, i) = "С���տ�ʱ��": i = i + 1
        .TextMatrix(0, i) = "�����տ���": i = i + 1
        .TextMatrix(0, i) = "�����տ�ʱ��": i = i + 1
        .TextMatrix(0, i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .ColData(0) = "-1|1"
       For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "�տ��" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "������" Or .ColKey(i) = "����ʱ��" Or _
               .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Then .ColData(i) = "1|0"
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ʵ���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False
    End With
End Sub
Private Function LoadHistoryData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʷ�տ�����
    '����:���ݼ��سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 17:08:50
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, i As Long, blnDel As Boolean
    Dim dtStartDate As Date, dtEndDate As Date
    On Error GoTo errHandle
    Call GetDateRange(dtStartDate, dtEndDate)
    
    If dtpEndDate - dtStartDate > 180 Then
        '���շ�Ա��Ϊͳ������
        Call MsgBox("�������õ�ʱ�䷶Χ�����˰���,�������Χ�Ĳ�ѯ", vbInformation + vbOKOnly, gstrSysName)
    End If
    If txtNO.Text <> "" Then
        strWhere = " And  A.NO = [3] "
    Else
        strWhere = " And  A.�Ǽ�ʱ�� Between [1] And [2] "
    End If
    
    strSQL = "" & _
    "   Select /*+ rule */a.Id,A.��¼����,decode(A.��¼����,4,'�����տ�',5,'�ֹ��տ�','���ʹ���') as ���,a.No As �տ��, a.�տ�Ա As �շ�Ա,  " & _
    "         b.���� As �տ��, a.ժҪ As �տ�˵��, " & _
    "         ltrim(to_char(a.��Ԥ����,'9999999999990.00')) as ��Ԥ����, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         a.�Ǽ��� as �����տ���,To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �����տ�ʱ��,  " & _
    "         a.������, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ�� " & _
    "  From ��Ա�սɼ�¼ A, ���ű� B " & _
    "  Where a.�տ��id = b.Id(+) And a.��¼����  in ( 4,5,6) " & strWhere & _
    "  Order by �Ǽ�ʱ�� desc,�տ�� desc,С���տ�ʱ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStartDate, dtEndDate, UCase(txtNO.Text))
    With vsCollect
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("��¼����")) = NVL(rsTemp!��¼����)
            .TextMatrix(.Rows - 1, .ColIndex("���")) = NVL(rsTemp!���)
            .TextMatrix(.Rows - 1, .ColIndex("�տ��")) = NVL(rsTemp!�տ��)
            .TextMatrix(.Rows - 1, .ColIndex("�շ�Ա")) = NVL(rsTemp!�շ�Ա)
            '.TextMatrix(.Rows - 1, .ColIndex("�տ��")) = Nvl(rsTemp!�տ��)
            .TextMatrix(.Rows - 1, .ColIndex("�տ�˵��")) = NVL(rsTemp!�տ�˵��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = NVL(rsTemp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ���")) = NVL(rsTemp!�����տ���)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ�ʱ��")) = NVL(rsTemp!�����տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTemp!����ʱ��)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then .Rows = .Rows - 1
'        If rsTemp.RecordCount <> 0 Then
'            Set .DataSource = rsTemp
'        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "�տ��" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "��¼����" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "�տ��" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("����ʱ��"))) <> ""
            If blnDel Then
                '���ϼ�¼���ú�ɫ����
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsCollect, Me.Name, "�տ���Ϣ�б�", False
        If .Enabled And .Visible Then .SetFocus
    End With
    mblnNotBrush = False
    '������ϸ����
    Call LoadDetail
    LoadHistoryData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2013-09-11 17:46:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngCollectID = 0
    mstrCollectNO = ""
    mblnDel = False: mint��¼���� = 4
    Call InitPage
    Call InitGrid '��ʼ������
    With cboDate
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 1: .ListIndex = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 2
        .AddItem "����"
        .ItemData(.NewIndex) = 3
        .AddItem "����"
        .ItemData(.NewIndex) = 4
        .AddItem "����"
        .ItemData(.NewIndex) = 5
        .AddItem "ָ��ʱ��"
        .ItemData(.NewIndex) = 9
    End With
    dtpEndDate.Value = zlDatabase.Currentdate
    dtpEndDate.MaxDate = Format(dtpEndDate.Value, "yyyy-mm-dd 23:59:59")
    dtpStartDate.Value = DateAdd("m", -1, dtpEndDate.Value)
    dtpStartDate.MaxDate = dtpEndDate.MaxDate
    Call SetCtrlVisible
End Sub
Private Sub SetCtrlVisible()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Visible����
    '����:���˺�
    '����:2013-09-11 18:21:29
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date, dtStartDate As Date
    dtpStartDate.Visible = False: dtpEndDate.Visible = False
    lblEndDate.Visible = False
    Select Case cboDate.ItemData(cboDate.ListIndex)
    Case 1 '����
        dtDate = zlDatabase.Currentdate
        lblRange.Caption = Format(dtDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 2 '����
        dtDate = DateAdd("d", -1, zlDatabase.Currentdate)
        lblRange.Caption = Format(dtDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 3 '����
        dtDate = zlDatabase.Currentdate
        dtStartDate = DateAdd("d", -1 * (Weekday(dtDate) - 2), dtDate)
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 4 '����
        dtDate = zlDatabase.Currentdate
        dtDate = DateAdd("d", -1 * (Weekday(dtDate) - 2), dtDate)
        dtStartDate = DateAdd("d", -7, dtDate)
        dtDate = DateAdd("d", 6, dtStartDate)
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 5 '����
        dtDate = zlDatabase.Currentdate
        dtStartDate = CDate(Format(dtDate, "yyyy") & "-" & Month(dtDate) & "-01")
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 9 'ָ������
        lblRange.Visible = False
        dtpStartDate.Visible = True: dtpEndDate.Visible = True
        lblEndDate.Visible = True
        If dtpStartDate.Enabled And dtpStartDate.Visible Then dtpStartDate.SetFocus
    End Select
End Sub
Private Function GetDateRange(ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡʱ�䷶Χ
    '���:dtStartDate-��ʼʱ��
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 18:45:57
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant
    On Error GoTo errHandle
    Select Case cboDate.ItemData(cboDate.ListIndex)
    Case 9 'ָ������
        dtStartDate = dtpStartDate.Value
        dtEndDate = dtpEndDate.Value
    Case Else '1, 2, 3, 4, 5 '���� '����'���� '����'����
        varData = Split(lblRange.Caption, "��")
        dtStartDate = CDate(varData(0))
        dtEndDate = CDate(varData(1))
    End Select
    GetDateRange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cboDate_Click()
    Call SetCtrlVisible
End Sub

Private Sub cmdRefresh_Click()
    Call LoadHistoryData
End Sub

Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub

Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Load()
    mblnDel = False
    Set mobjChargeBill = New clsChargeBill
    Call InitPanel
    Call InitFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "�տ���Ϣ�б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
    Set mobjChargeBill = Nothing
End Sub
 
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
        tbPage.Width = .ScaleWidth
    End With
End Sub

Private Sub picRollingCurtain_Resize()
    Err = 0: On Error Resume Next
    With picRollingCurtain
        vsRollingCurtain.Left = .ScaleLeft
        vsRollingCurtain.Top = .ScaleTop
        vsRollingCurtain.Height = .ScaleHeight
        vsRollingCurtain.Width = .ScaleWidth
    End With
End Sub
Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
    zlCommFun.OpenIme False
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Or Trim(txtNO.Text) = "" Then Exit Sub
    txtNO.Text = GetFullNO(Trim(txtNO.Text), 141)
    Call LoadHistoryData
End Sub

Private Sub vsCollect_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
End Sub

Private Sub vsRollingCurtain_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
    vsRollingCurtain.Tag = "0"
End Sub

Private Sub vsRollingCurtain_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub picCollect_Resize()
    Err = 0: On Error Resume Next
    Line1.X2 = picCollect.Width
    With picCollect
        vsCollect.Left = .ScaleLeft
        vsCollect.Top = txtNO.Top + txtNO.Height + 100
        vsCollect.Height = .ScaleHeight - vsCollect.Top - 50
        vsCollect.Width = .ScaleWidth
    End With
End Sub
Private Sub vsCollect_GotFocus()
    Call zl_VsGridGotFocus(vsCollect)
    vsCollect.Tag = "1"
End Sub
Private Sub vsCollect_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsCollect)
    vsCollect.Tag = "0"
End Sub
Private Sub vsCollect_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "�տ���Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsCollect_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsCollect, OldRow, NewRow, OldCol, NewCol)
    'If OldRow = NewRow Or mblnNotBrush Then Exit Sub
    Call LoadDetail
End Sub
Private Sub vsCollect_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "�տ���Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Function LoadDetail() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϸ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 11:17:09
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date, strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long
    Dim lng�տ�ID As Long, strNO As String, blnDel As Boolean
    With vsCollect
        If .Row >= 1 Then
            lng�տ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strNO = Trim(.TextMatrix(.Row, .ColIndex("�տ��")))
            blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
             mint��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
        End If
    End With
    mlngCollectID = lng�տ�ID
    mstrCollectNO = strNO
    mblnDel = blnDel
    
    On Error GoTo errHandle
    dtStartDate = CDate("1991-01-01"): dtEndDate = dtStartDate
    
    If Val(tbPage.Selected.Tag) = mPgIndex.EM_PG_�տƱ�� Then
        '����Ʊ�ݻ���
        If lng�տ�ID = 0 Then
            mobjChargeBill.ClearChargeAndBillTotalForm
        Else
            If mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 4, lng�տ�ID, dtStartDate, dtEndDate, True, blnDel) = False Then Exit Function
        End If
        LoadDetail = True
        'Exit Function
    End If
    '����������Ϣ�б�
    strSQL = "" & _
    "   Select /*+ rule */a.Id,a.No As ���ʵ���,Substr(Decode(�Ƿ�Һ�,1,',�Һ�','') || Decode(�Ƿ���￨,1,',���￨','') || Decode(�Ƿ����ѿ�,1,',���ѿ�','') || Decode(�Ƿ��շ�,1,',�շ�','') || Decode(�Ƿ����,1,',����','') || Decode(Ԥ�����,1,',Ԥ��',2,',����Ԥ��',3,',סԺԤ��',''),2) As �������," & _
    "         a.��ʼʱ��, a.��ֹʱ��, a.�Ǽ��� As ������, a.�Ǽ�ʱ�� As ����ʱ��,  " & _
    "         b.���� As �տ��, a.ժҪ As ����˵��, " & _
    "         ltrim(to_char(a.��Ԥ����,'9999999999990.00')) as ��Ԥ����, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         a.С���տ���, To_Char(a.С���տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As С���տ�ʱ��,  " & _
    "         a.�����տ���,To_Char(a.�����տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �����տ�ʱ��,  " & _
    "         a.������, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ�� " & _
    "  From ��Ա�սɼ�¼ A, ���ű� B " & _
    "  Where a.�տ��id = b.Id(+) And a.��¼����  in (3,1) and A.�����տ�ID= [1]" & _
    "  Order by С������ID Desc,�Ǽ�ʱ�� desc,���ʵ��� desc,С���տ�ʱ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�տ�ID)
    
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .OutlineBar = flexOutlineBarComplete
        .FixedRows = 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = NVL(rsTemp!���ʵ���)
            .TextMatrix(.Rows - 1, .ColIndex("�������")) = NVL(rsTemp!�������)
            .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = NVL(rsTemp!��ʼʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = NVL(rsTemp!��ֹʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTemp!����ʱ��)
            '.TextMatrix(.Rows - 1, .ColIndex("�տ��")) = Nvl(rsTemp!�տ��)
            .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = NVL(rsTemp!����˵��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = NVL(rsTemp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ���")) = NVL(rsTemp!С���տ���)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ�ʱ��")) = NVL(rsTemp!С���տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ���")) = NVL(rsTemp!�����տ���)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ�ʱ��")) = NVL(rsTemp!�����տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTemp!����ʱ��)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then .Rows = .Rows - 1
'        If rsTemp.RecordCount <> 0 Then
'            Set .DataSource = rsTemp
'        End If
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "�տ��" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) = "���ʵ���" Then
                .OutlineCol = i
            End If
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ʵ���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("����ʱ��"))) <> ""
            If blnDel Then
                '���ϼ�¼���ú�ɫ����
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            If Val(.TextMatrix(i, .ColIndex("ID"))) <> 0 Then
                If .TextMatrix(i, .ColIndex("С���տ�ʱ��")) = "" Then
                    .IsSubtotal(i) = True
                Else
                    .RowOutlineLevel(i) = 1
                End If
            End If
        Next
        .Outline (0)
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False
        If .Enabled And .Visible Then .SetFocus
    End With
    
    LoadDetail = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Property Get GetChargeCollectID() As Long
    GetChargeCollectID = mlngCollectID
End Property

Public Property Get GetChargeCollectNO() As String
    GetChargeCollectNO = mstrCollectNO
End Property
 
Private Function CheckCancelValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������տ����ݵĺϷ���
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 15:44:57
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�տ�ID As Long, strNO As String, blnDel As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, strPerson As String
    Dim int��¼���� As Integer
    On Error GoTo errHandle
    With vsCollect
        lng�տ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("�տ��")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
        strPerson = Trim(.TextMatrix(.Row, .ColIndex("�շ�Ա")))
        int��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
    End With
    
    If blnDel Then
        MsgBox "�տ��Ϊ:" & strNO & "���տ��¼�Ѿ������ϣ�������������!", vbInformation + vbOKOnly, gstrSysName
        If vsCollect.Enabled And vsCollect.Visible Then vsCollect.SetFocus
        Exit Function
    End If
    
    If int��¼���� = 4 Then
        '����Ƿ����һ���տ�
        strSQL = "" & _
        "   Select 1 From ��Ա�սɼ�¼  " & _
        "   Where �Ǽ�ʱ��>(Select Max(�Ǽ�ʱ��) From ��Ա�սɼ�¼ where ID=[1] ) " & _
        "               And ID+0<>[1] AND Rownum <2 And ��¼����=4 And �տ�Ա=[2] and ����ʱ�� is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�տ�ID, strPerson)
        If rsTemp.EOF = False Then
            MsgBox "ע��:" & vbCrLf & _
                            "       �տ��Ϊ:" & strNO & "���տ��¼���������һ��" & _
                           "���տ��¼,Ϊ�˱�֤�տ�������ȷ�����������һ���տ��¼��ʼ����!", vbInformation + vbOKOnly, gstrSysName
            If vsCollect.Enabled And vsCollect.Visible Then vsCollect.SetFocus
            Exit Function
        End If
    End If
    CheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function CancelData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϵ�ǰ�տ�����
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 15:44:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�տ�ID As Long, strNO As String, blnDel As Boolean
    Dim strDate As String, strSQL As String
    
    On Error GoTo errHandle
    With vsCollect
        If .Row < 1 Then Exit Function
        If .ColIndex("�տ��") < 0 _
            Or .ColIndex("ID") < 0 _
            Or .ColIndex("����ʱ��") < 0 Then Exit Function
        lng�տ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("�տ��")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
        mint��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
        If strNO = "" Then Exit Function
    End With
    If CheckCancelValied = False Then Exit Function
    
    If MsgBox("���Ƿ����Ҫ���տ��Ϊ:" & strNO & "����������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'Zl_�����տ��¼_Cancel
    strSQL = "Zl_�����տ��¼_Cancel("
    '  Id_In       In ��Ա�սɼ�¼.Id%Type,
    strSQL = strSQL & "" & lng�տ�ID & ","
    ' ��¼����_In In ��Ա�սɼ�¼.��¼����%Type,
    strSQL = strSQL & "" & mint��¼���� & ","
    '  ������_In   In ��Ա�սɼ�¼.������%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In In ��Ա�սɼ�¼.����ʱ��%Type
    strSQL = strSQL & "to_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    With vsCollect
        .TextMatrix(.Row, .ColIndex("������")) = UserInfo.����
        .TextMatrix(.Row, .ColIndex("����ʱ��")) = strDate
        .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
         mblnDel = True
    End With
    CancelData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String, blnPrintRollingCurtain As Boolean
    Dim rsTemp As ADODB.Recordset, lng�տ�ID As Long, strNO As String, blnDel As Boolean
    Err = 0: On Error GoTo ErrHand:
    blnPrintRollingCurtain = False
    If Val(vsCollect.Tag) = 0 Then
        '��ӡ�տƱ�ݻ���
        With vsCollect
            If .Row < 1 Then Exit Sub
            If .ColIndex("�տ��") < 0 _
                  Or .ColIndex("ID") < 0 _
                  Or .ColIndex("����ʱ��") < 0 Then Exit Sub
              lng�տ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
              strNO = Val(.TextMatrix(.Row, .ColIndex("�տ��")))
              blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
              If lng�տ�ID = 0 Then Exit Sub
        End With
        If Val(tbPage.Selected.Tag) = mPgIndex.EM_PG_�տƱ�� Then
            Call mobjChargeBill.zlPrint(bytMode): Exit Sub
        End If
        '��ӡ������Ϣ
        blnPrintRollingCurtain = True
    End If
    
    '����տ���Ϣ
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr��λ���� & "�����տ����"
    Set objRow = New zlTabAppRow
    If blnPrintRollingCurtain Then
        objRow.Add "�տ�ţ�" & strNO
    Else
        If lblRange.Visible Then
            objRow.Add "ʱ�䷶Χ��" & lblRange.Caption
        Else
            objRow.Add "ʱ�䷶Χ��" & Format(dtpStartDate, "yyyy-mm-dd HH:MM:SS") & "��" & Format(dtpEndDate, "yyyy-mm-dd HH:MM:SS")
        End If
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = IIf(blnPrintRollingCurtain, vsRollingCurtain, vsCollect)
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش��տ��վ�
    '����:���˺�
    '����:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, int��¼���� As Integer
    If Not (zlStr.IsHavePrivs(mstrPrivs, "�տ��վݴ�ӡ") And zlStr.IsHavePrivs(mstrPrivs, "�ش��տ��վ�")) Then Exit Sub
    With vsCollect
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("�տ��")))
        int��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
    End With
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "��¼����=" & int��¼����, 2)
End Sub

Public Sub zlRefresh()
    '���½�������ˢ��
    Call cmdRefresh_Click
End Sub

Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ�տ�����
    '����:���˺�
    '����:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�տ�ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    
    With vsCollect
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("�տ��")))
        lng�տ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
    End With
    
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
   Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, 4, CStr(lng�տ�ID), dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
End Sub

Public Property Get IsAllowViewChargeList() As Boolean
    '�Ƿ�����鿴��ϸ
    Dim int��¼���� As Integer, lngID As Long
    With vsCollect
        If .Row < 1 Then Exit Property
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Property
        int��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
        IsAllowViewChargeList = int��¼���� = 4
    End With
End Property

Public Property Get IsAllowCollectCancel() As Boolean
    '�Ƿ������տ�����
    Dim int��¼���� As Integer, lngID As Long
    With vsCollect
        If .Row < 1 Then Exit Property
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Property
        int��¼���� = Val(.TextMatrix(.Row, .ColIndex("��¼����")))
        IsAllowCollectCancel = int��¼���� <> 6 And mblnDel = False
    End With
End Property

Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long
    Dim lng�տ�ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsCollect
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("�տ��")))
        lng�տ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "�տ��=" & strNO, _
        "�տ�ID=" & lng�տ�ID, _
        "���ϱ�־=" & IIf(blnDel, 1, 0))
End Sub

Private Sub imgColPlanCollect_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanCollect.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + imgColPlanCollect.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollect, lngLeft, lngTop, imgColPlanCollect.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollect, Me.Name, "�տ���Ϣ�б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlanCollect_Click()
    Call imgColPlanCollect_Click
End Sub

Private Sub imgColPlanRC_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanRC.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + imgColPlanRC.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlanRC.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlanRC_Click()
    Call imgColPlanRC_Click
End Sub




