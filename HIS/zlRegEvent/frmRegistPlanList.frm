VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmRegistPlanList 
   BorderStyle     =   0  'None
   Caption         =   "��ǰ��Ч�ű�"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTime 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   2295
      ScaleHeight     =   2745
      ScaleWidth      =   4425
      TabIndex        =   7
      Top             =   4485
      Width           =   4425
      Begin MSComctlLib.TabStrip tbWeekTime 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   60
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   503
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTime 
         Height          =   2145
         Left            =   0
         TabIndex        =   8
         Top             =   375
         Width           =   4410
         _cx             =   7779
         _cy             =   3784
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegistPlanList.frx":0000
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
         ExplorerBar     =   0
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
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   -2055
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   5
      Top             =   4170
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2295
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4575
         _Version        =   589884
         _ExtentX        =   8070
         _ExtentY        =   4048
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picStop 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   4455
      ScaleHeight     =   3285
      ScaleWidth      =   3660
      TabIndex        =   3
      Top             =   3840
      Width           =   3660
      Begin VSFlex8Ctl.VSFlexGrid vsStop 
         Height          =   2415
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   6345
         _cx             =   11192
         _cy             =   4260
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanList.frx":0075
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
         ExplorerBar     =   7
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   75
      ScaleHeight     =   2250
      ScaleWidth      =   5160
      TabIndex        =   0
      Top             =   750
      Width           =   5160
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2145
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3405
         _cx             =   6006
         _cy             =   3784
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   26
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanList.frx":0164
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
         ExplorerBar     =   7
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   2
            Top             =   210
            Width           =   210
            Begin VB.Image imgColList 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistPlanList.frx":04C5
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
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
Attribute VB_Name = "frmRegistPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mArrFilter As Variant  '��������
Private mblnNotMoveRow As Boolean '���ƶ���
Private mblnListSetFocus As Boolean  '��ǰ����λ��
Private mrsRoom As ADODB.Recordset
Private mrsPlanRoom As ADODB.Recordset
Private Const conPane_List = 1
Private Const conPane_Plan = 2
Private Const conPane_Stop = 3
Private mblnHaveDate As Boolean
Private mblnHaveUnit As Boolean
Public Event zlPopuMenu(intType As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Private WithEvents mfrm�ƻ� As frmRegistPlanIntend
Attribute mfrm�ƻ�.VB_VarHelpID = -1

Private Enum mPgIndex
    pg_�ƻ� = 1
    pg_ʱ�� = 2
    pg_ͣ�� = 3
    pg_��λ = 4
End Enum
Private mblnNotClick As Boolean
Private mrsTime As ADODB.Recordset
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(conPane_List, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "�ҺŰ�����Ϣ"
    panThis.Handle = picList.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = conPane_List
    Set panThis = dkpMan.CreatePane(conPane_Plan, 250, 580, DockBottomOf, panThis)
    panThis.Title = ""
    panThis.Tag = conPane_Plan
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picPage.Hwnd
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = picList.Hwnd
    Case conPane_Plan
        Item.Handle = picPage.Hwnd
    End Select
End Sub
 
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer, i As Integer, objGrid As VSFlexGrid
    i = 0
    With vsList
        .Redraw = flexRDNone
        .Rows = 3: .FixedRows = 2
        .FixedCols = 1
        .Cols = 39:   .Clear
        .FrozenCols = 6
        .TextMatrix(0, i) = "  ": .ColWidth(i) = 285
        .TextMatrix(1, i) = "  ":  .ColKey(i) = "��־": i = i + 1
        
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0
        .TextMatrix(1, i) = "ID": .ColKey(i) = "ID": i = i + 1
         
        .TextMatrix(0, i) = "״̬": .ColWidth(i) = 200
        .TextMatrix(1, i) = "״̬": .ColKey(i) = "״̬": i = i + 1
         
        .TextMatrix(0, i) = "����": .ColWidth(i) = 720
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1

        .TextMatrix(0, i) = "�ű�": .ColWidth(i) = 480
        .TextMatrix(1, i) = "�ű�": .ColKey(i) = "�ű�": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1
        .TextMatrix(0, i) = "��Ŀ": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "��Ŀ": .ColKey(i) = "��Ŀ": i = i + 1
        .TextMatrix(0, i) = "ҽ��":: .ColWidth(i) = 1000
        .TextMatrix(1, i) = "ҽ��": .ColKey(i) = "ҽ��": i = i + 1
        .TextMatrix(0, i) = "�ѹ���": .ColWidth(i) = 660
        .TextMatrix(1, i) = "�ѹ���": .ColKey(i) = "�ѹ���": i = i + 1
        .TextMatrix(0, i) = "��Լ��": .ColWidth(i) = 660
        .TextMatrix(1, i) = "��Լ��": .ColKey(i) = "��Լ��": i = i + 1
        '59611 ���ϴ� 2014/04/11 16:40:49 ���һ����ʾԤԼ�ѽ�����
        .TextMatrix(0, i) = "�����ѽ�����": .ColWidth(i) = 1305
        .TextMatrix(1, i) = "�����ѽ�����": .ColKey(i) = "�����ѽ�����": i = i + 1
        
        .TextMatrix(0, i) = "����": .ColWidth(i) = 495
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "��һ": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "��һ-����": i = i + 1
        .TextMatrix(0, i) = "��һ": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "��һ-�޺�": i = i + 1
        .TextMatrix(0, i) = "��һ": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "��һ-��Լ": i = i + 1

        .TextMatrix(0, i) = "�ܶ�": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "�ܶ�-����": i = i + 1
        .TextMatrix(0, i) = "�ܶ�": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "�ܶ�-�޺�": i = i + 1
        .TextMatrix(0, i) = "�ܶ�": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "�ܶ�-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 420
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1
        .TextMatrix(0, i) = "���﷽ʽ": .ColWidth(i) = 855
        .TextMatrix(1, i) = "���﷽ʽ": .ColKey(i) = "���﷽ʽ": i = i + 1
        .TextMatrix(0, i) = "IDS": .ColWidth(i) = 0: .ColHidden(i) = True
        .TextMatrix(1, i) = "IDS": .ColKey(i) = "IDS": i = i + 1
        .TextMatrix(0, i) = "��Ч��Χ": .ColWidth(i) = 2800
        .TextMatrix(1, i) = "��Ч��Χ": .ColKey(i) = "��Ч��Χ": i = i + 1
        .TextMatrix(0, i) = "��ſ���": .ColWidth(i) = 765
        .TextMatrix(1, i) = "��ſ���": .ColKey(i) = "��ſ���": i = i + 1
        .TextMatrix(0, i) = "ͣ������": .ColWidth(i) = 1860
        .TextMatrix(1, i) = "ͣ������": .ColKey(i) = "ͣ������": i = i + 1
        .TextMatrix(0, i) = "Ӧ������": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "Ӧ������": .ColKey(i) = "Ӧ������": i = i + 1
        .Cell(flexcpText, 0, 0, .Rows - 1) = " "
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        For i = 0 To .Cols - 1
            .MergeCol(i) = True:
            .FixedAlignment(i) = flexAlignCenterCenter
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case .ColKey(i)
            Case "ID", "��־", "IDS"
                 .ColData(i) = "-1|1"
            Case "����", "�ű�"
                .ColData(i) = "1|0"
            End Select
        Next
         .MergeRow(0) = True: .MergeRow(1) = True
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    Call InitPancel
    Call InitUnitData
    Call InitPage
    Call InitVsGrid
    Call mfrm�ƻ�.SetGotFocus(True): Call vsList_LostFocus: Call vsStop_LostFocus
    vsList_GotFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "��Ч��-�ű��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
    zl_vsGrid_Para_Save mlngModule, vsStop, Me.Caption, "��Ч��-ͣ�üƻ�", True, , InStr(1, mstrPrivs, ";��������;") > 0
    If Not mfrm�ƻ� Is Nothing Then Unload mfrm�ƻ�
    Unload frmUnitReg
    Set mfrm�ƻ� = Nothing
End Sub

Public Sub ReloagUnitRegPlan()
    If mfrm�ƻ� Is Nothing Then Exit Sub
    mfrm�ƻ�.ReLoadUnitPlan
End Sub


Private Sub InitUnitData()
    '��ʼ��������λ��Ϣ
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    strSQL = "Select Count(0) as count From �Һź�����λ Where Rownum=1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mblnHaveUnit = Val(Nvl(rsTmp!Count)) > 0
    Set rsTmp = Nothing
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub imgColList_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgList.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsList, lngLeft, lngTop, imgColList.Height)
    
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "��Ч��-�ű��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
 
 
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Left = .ScaleLeft
        vsList.Width = .ScaleWidth
        vsList.Top = .ScaleTop
        vsList.Height = .ScaleHeight
    End With
End Sub

 

Private Sub picPage_Resize()
    Err = 0: On Error Resume Next
    With picPage
        tbPage.Left = .ScaleLeft
        tbPage.Width = .ScaleWidth
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub picTime_Resize()
    Err = 0: On Error Resume Next
    With picTime
        tbWeekTime.Left = .ScaleLeft + 100
        tbWeekTime.Top = .ScaleTop + 50
        tbWeekTime.Width = .ScaleWidth
        If tbWeekTime.Visible = False Then
            vsTime.Top = .ScaleTop
        Else
            vsTime.Top = tbWeekTime.Top + tbWeekTime.Height + 50
        End If
        vsTime.Left = .ScaleLeft
        vsTime.Width = .ScaleWidth
       ' vsTime.Top = .ScaleTop
        vsTime.Height = .ScaleHeight - vsTime.Top
    End With
End Sub
 

Private Function HaveData() As Boolean
    '����:�Ƿ�������
    If Me.ActiveControl Is vsList Then
        With Me.vsList
            HaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    Else
        HaveData = mfrm�ƻ�.zlHaveData
    End If
End Function

Public Function Have�ƻ�() As Boolean
    '����:�ð������Ƿ��мƻ���Ϣ
    '�����:51429
    Dim i As Long
    Dim blnHasData As Boolean
    
    For i = 0 To mfrm�ƻ�.vsPlan.Rows - 1
       With mfrm�ƻ�.vsPlan
           If Val(Nvl(.TextMatrix(i, .ColIndex("ID")), "0")) <> 0 Then
                blnHasData = True
                Exit For
           End If
       End With
    Next
    Have�ƻ� = blnHasData
End Function
Public Function �Ƿ�ѡ�мƻ��б�() As Boolean
    '����:�ð������Ƿ�ѡ�мƻ��б�
    '�����:51429
    �Ƿ�ѡ�мƻ��б� = mfrm�ƻ�.mblnSelected�ƻ�
End Function
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid, rsTemp As New ADODB.Recordset, strSQL As String
     If Not Me.ActiveControl Is vsList Then
        mfrm�ƻ�.zlRptPrint (bytFunc): Exit Sub
     End If
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & IIf(Not Me.ActiveControl Is vsList, "�ҺŰ��ű�", "�Һżƻ���")
    
    If CStr(mArrFilter("��Ч��")(0)) <> "1901-01-01" Then
        objRow.Add "Ч�ڷ�Χ��" & CStr(mArrFilter("��Ч��")(0)) & "��" & CStr(mArrFilter("��Ч��")(1))
    End If
    If Val(mArrFilter("����ID")) > 0 Then
        strSQL = "Select ���� From ���ű� where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("����ID")))
        If rsTemp.EOF Then
            objRow.Add "���ң����п���"
        Else
            objRow.Add "���ң�" & Nvl(rsTemp!����)
        End If
    ElseIf Val(mArrFilter("����ID")) = -1 Then
        objRow.Add "���ң�����Ա��������"
    Else
        objRow.Add "���ң����п���"
    End If
    Select Case mArrFilter("ҽ��ID")(1)
    Case "ID"
        strSQL = "Select ���� From ��Ա�� where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("ҽ��ID")(0)))
        If rsTemp.EOF Then
            objRow.Add "ҽ��������"
        Else
            objRow.Add "ҽ����" & Nvl(rsTemp!����)
        End If
    Case "UPR", "NONE"
            objRow.Add "ҽ����" & CStr(mArrFilter("ҽ��ID")(0))
    End Select
    objPrint.UnderAppRows.Add objRow
    Set vsGrid = vsList
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("��־") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
Private Sub picStop_Resize()
    Err = 0: On Error Resume Next
    With picStop
        vsStop.Left = .ScaleLeft
        vsStop.Top = .ScaleTop
        vsStop.Height = .ScaleHeight
        vsStop.Width = .ScaleWidth
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call picTime_Resize
End Sub

Private Sub tbWeekTime_Click()
    Dim lng����ID  As Long, bln��ſ��� As Boolean
    If mblnNotClick = True Then Exit Sub
   With vsList
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln��ſ��� = Trim(.TextMatrix(.Row, .ColIndex("��ſ���"))) <> ""
    End With
    Call LoadTimePlan(lng����ID, bln��ſ���)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng����ID As Long, bln��ſ��� As Boolean
    Dim str���� As String, str�ű� As String, str���� As String
    If OldRow > 0 Then
        zl_VsGridRowChange vsList, OldRow, NewRow, OldCol, NewCol
    End If
    
    If OldRow = NewRow Or mblnNotMoveRow Then Exit Sub
    
    With vsList
        lng����ID = Val(.TextMatrix(NewRow, .ColIndex("ID")))
        bln��ſ��� = Trim(.TextMatrix(NewRow, .ColIndex("��ſ���"))) <> ""
        str���� = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        str�ű� = Trim(.TextMatrix(NewRow, .ColIndex("�ű�")))
        str���� = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        
    End With
    Call mfrm�ƻ�.zlShowPlan(lng����ID, str����, str�ű�, str����)
    Call LoadStopPlan(lng����ID)
    Call LoadTimePlan(lng����ID, bln��ſ���)
    Call LoadUnitReg(lng����ID)
    If OldRow <> NewRow Then
        zlHaveDatPlanForPlan
    End If
    '�����:51429
    mfrm�ƻ�.mblnSelected�ƻ� = False
    On Error Resume Next
    If vsList.Enabled And vsList.Visible Then vsList.SetFocus
    
    DoEvents
    Call vsList_GotFocus
End Sub

Private Sub LoadUnitReg(ByVal lng����ID As Long)
     frmUnitReg.ShowUnitReg lng����ID
End Sub

Public Sub zlHaveDatPlanForPlan()
    Dim lngNum  As Long
    Dim lngRow  As Long
    
     With vsList
        If .Row = 0 Then Exit Sub
        lngRow = .Row
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("����-�޺�")))
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("��һ-�޺�"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("�ܶ�-�޺�"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("����-�޺�"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("����-�޺�"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("����-�޺�"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("����-�޺�"))) + lngNum
     End With
     mblnHaveDate = lngNum > 0
End Sub

Public Sub ReloadTimePlan(Optional ByVal blnReloadPlan As Boolean = False)
    '***********************************
    '�ҺŰ���ʱ�θ����Ժ�
    '���¹ҺŰ���ʱ����ʾ�б�
    '***********************************
    Dim lng����ID       As Long
    Dim bln��ſ���     As Boolean
    If blnReloadPlan Then
        mfrm�ƻ�.ReloadTimePlan
        Exit Sub
    End If
    With vsList
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln��ſ��� = Trim(.TextMatrix(.Row, .ColIndex("��ſ���"))) <> ""
    End With
    Call LoadTimePlan(lng����ID, bln��ſ���, False, True)
    zlControl.ControlSetFocus vsList, True
End Sub
Private Sub LoadStopPlan(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͣ�üƻ�����
    '����:���˺�
    '����:2010-09-09 11:54:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, strCurDate As String
    On Error GoTo errHandle
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM")
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    strSQL = "Select ����ID,���,��ʼֹͣʱ��,����ֹͣʱ��,�ƶ���,�ƶ�����,��ע From �ҺŰ���ͣ��״̬ where ����ID=[1] Order by ��ʼֹͣʱ��,�ƶ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    With vsStop
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("���")) = i
            .Cell(flexcpData, i, .ColIndex("���")) = Val(Nvl(rsTemp!���))
            .TextMatrix(i, .ColIndex("��ʼͣ��ʱ��")) = Format(rsTemp!��ʼֹͣʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("����ͣ��ʱ��")) = Format(rsTemp!����ֹͣʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("�ƶ���")) = Nvl(rsTemp!�ƶ���)
            .TextMatrix(i, .ColIndex("�ƶ�����")) = Format(rsTemp!�ƶ�����, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("��ע")) = Nvl(rsTemp!��ע)
            If Format(rsTemp!����ֹͣʱ��, "yyyy-mm-dd HH:MM:SS") < strCurDate Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H8000000C
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        '�ָ�������
         zl_vsGrid_Para_Restore mlngModule, vsStop, Me.Caption, "��Ч��-ͣ�üƻ�", True
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadTimePlan(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False, Optional blnReload As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͣ�üƻ�����
    '���:bln�ƻ�-�Ƿ���ؼƻ���ʱ���
    '����:���˺�
    '����:2010-09-09 11:54:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str���� As String
    Dim i As Long, r As Integer, strʱ�� As String, strTime As String, strKey As String
    Static lngPre����ID  As Long
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
        lngPre����ID = -1
    ElseIf mrsTime.State <> 1 Then
         lngPre����ID = -1
    End If
    If lngPre����ID <> lng����ID Or blnReload Then
        lngPre����ID = lng����ID
        strSQL = "" & _
        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               ��������,�Ƿ�ԤԼ" & _
        "   From  �ҺŰ���ʱ��  " & _
        "   Where ����ID=[1] And ��������>0" & _
        "   Order by ����,ʱ��,���"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!����) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
                    strTime = Nvl(mrsTime!����)
                End If
                .MoveNext
            Loop
            mblnNotClick = True
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            mblnNotClick = False
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        
            Call picTime_Resize
        End With
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 800: .RowHeightMin = 800
        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             r = 0: i = -1
            Do While Not mrsTime.EOF
                i = i + 1
                If i > .Cols - 1 Then r = r + 1: i = 0
                strTime = "ԤԼ" & Val(Nvl(mrsTime!��������)) & "��" & vbCrLf & vbCrLf
                strTime = strTime & mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i) = strTime
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
                r = r + 1
                strʱ�� = Nvl(mrsTime!ʱ��)
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, 0) = strʱ��
                i = 0
            End If
            i = i + 1
            strTime = mrsTime!��� & vbCrLf & vbCrLf
            strTime = strTime & mrsTime!ʱ�䷶Χ
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "��Ч��-�ű��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
        With vsList
            If .ColKey(Col) Like "��*" Then
                 Position = Col
            End If
        End With
End Sub

Private Sub vsList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent zlPopuMenu(0, Button, Shift, X, Y)
End Sub
    
Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsList
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub

Private Sub vsList_DblClick()
    With vsList
        If Trim(.TextMatrix(.Row, .ColIndex("ͣ������"))) <> "" Then
            Call frmRegistPlanEdit.ShowEdit(Me, edt_����, mlngModule, mstrPrivs, .TextMatrix(.Row, .ColIndex("id")))
            Exit Sub
        End If
    End With
    zlExecuteModifyList Me
End Sub
 
Private Sub vsList_GotFocus()
    zl_VsGridGotFocus vsList
    mblnListSetFocus = True
End Sub
Private Sub vsList_LostFocus()
    zl_VsGridLOSTFOCUS vsList
    With vsList
        .ForeColorSel = .Cell(flexcpForeColor, .Row, .Col)
    End With
End Sub
 
Private Sub LoadDataToList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long, strSQL As String
    Dim blnHistory As Boolean, strStartDate As String, rs�ҺŻ��� As ADODB.Recordset, strWhere As String
    Dim lngPreID As Long, strTable As String
    Dim lng�������� As Long
    Dim str��ʾ���� As String
    
    Err = 0: On Error GoTo Errhand:
    strWhere = ""
    If CStr(mArrFilter("��Ч��")(0)) <> "1901-01-01" Then
        strFilter = "  And Nvl(A.��ʼʱ��,To_Date('3000-01-01','YYYY-MM-DD'))>=[4]   And Nvl(A.��ֹʱ��,To_Date('1900-01-01','YYYY-MM-DD'))<=[5]"
    End If
    If Val(mArrFilter("����ID")) > 0 Then
        strFilter = strFilter & " And A.����ID=[1]"
        strWhere = strWhere & " And A.����ID=[1]"
    End If
    If Val(mArrFilter("����ID")) = -1 Then
        strFilter = strFilter & " And  A.����ID in (Select ����ID From ������Ա where ��Աid=[6]) "
        strWhere = strWhere & " And  A.����ID in (Select ����ID From ������Ա where ��Աid=[6]) "
    End If
    Select Case mArrFilter("ҽ��ID")(1)
    Case "ID"
         strFilter = strFilter & "  And A.ҽ��ID=[2]"
         strWhere = strWhere & "  And A.ҽ��ID=[2]"
    Case "UPR"
         strFilter = strFilter & " And Upper(A.ҽ������)=[3]"
         strWhere = strWhere & " And Upper(A.ҽ������)=[3]"
    Case "NONE"
         strFilter = strFilter & " And A.ҽ������=[3]"
         strWhere = strWhere & " And A.ҽ������=[3]"
    End Select
    strTable = "" & _
    "   Select A.ID, " & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'��һ',B.�޺���,0)) as ��һ�޺�, Sum(Decode(B.������Ŀ,'��һ',B.��Լ��))  as ��һ��Լ," & _
    "             Sum(Decode(B.������Ŀ,'�ܶ�',B.�޺���,0)) as �ܶ��޺�, Sum(Decode(B.������Ŀ,'�ܶ�',B.��Լ��))  as �ܶ���Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ" & _
    "   From �ҺŰ��� A,�ҺŰ������� B  " & _
    "   Where A.ID=B.����ID(+) " & strFilter & _
    "   Group by A.ID"
    '����:32512��38505
    strSQL = "" & _
    "   Select ����id, ��Ŀid,  nvl(ҽ������,'���˺����ҽ��') As ҽ������, nvl(ҽ��id,0) As ҽ��ID, A.����, �ѹ���, ��Լ��, �����ѽ��� " & _
    "   From ���˹ҺŻ��� A " & _
    "   Where ���� = Trunc(Sysdate) " & strWhere
     '���� :45525
    If Nvl(mArrFilter("��ʾͣ�ð���"), 0) = 0 And Nvl(mArrFilter("��ʾɾ������"), 0) = 0 Then 'ͣ����ɾ��ͬʱ����ʾ��ʱ��
        str��ʾ���� = "And (Nvl(a.ͣ������,to_date('3000-1-1','yyyy-mm-dd')) > sysdate)"
    ElseIf Nvl(mArrFilter("��ʾͣ�ð���"), 0) = 1 And Nvl(mArrFilter("��ʾɾ������"), 0) = 0 Then 'ͣ����ʾ��ɾ������ʾ��ʱ��
        str��ʾ���� = "And (A.�Ƿ�ɾ��<> 1)"
    ElseIf Nvl(mArrFilter("��ʾͣ�ð���"), 0) = 0 And Nvl(mArrFilter("��ʾɾ������"), 0) = 1 Then 'ͣ�ò���ʾ��ɾ����ʾ��ʱ��
        str��ʾ���� = "And(A.�Ƿ�ɾ��=1or Nvl(a.ͣ������,to_date('3000-1-1','yyyy-mm-dd')) > sysdate)"
    End If
    
    strSQL = _
      "Select A.ID,A.���� as �ű�,A.����,A.����ID,C.���� as ����,A.��ĿID,B.���� as ��Ŀ," & _
      "         A.ҽ������ as ҽ��,A.ҽ��ID, Nvl(A.��������,0) as ����,Nvl(A.��ſ���,0) as ��ſ���," & _
      "         Decode(Nvl(A.���﷽ʽ,0),0,'������',1,'ָ������',2,'��̬����',3,'ƽ������') as ���﷽ʽ," & _
      "         A.����,A1.�����޺�,A1.������Լ,A.��һ,A1.��һ�޺�,A1.��һ��Լ,A.�ܶ�,A1.�ܶ��޺�,A1.�ܶ���Լ, " & _
      "         A.����,A1.�����޺�,A1.������Լ,A.����,A1.�����޺�,A1.������Լ,A.����,A1.�����޺�,A1.������Լ, " & _
      "         A.����,A1.�����޺�,A1.������Լ, " & _
      "         A.��ʼʱ��,A.��ֹʱ��, " & _
      "         D.�ѹ���,D.��Լ��,D.�����ѽ���,to_char(A.ͣ������,'yyyy-mm-dd HH24:mi:ss') as ͣ������,a.�Ƿ�ɾ��" & _
      " From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� C," & vbCrLf & _
      "         (" & strTable & ") A1, " & _
      "         (" & strSQL & ") D" & _
      " Where A.ID=A1.ID and a.��ĿID = B.ID And a.����ID = C.ID And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & _
      "            And A.����ID=D.����ID(+) and A.��ĿID=D.��ĿID(+) " & _
      "             And nvl(A.ҽ������,'���˺����ҽ��')= D.ҽ������(+) And nvl(A.ҽ��id,0)=D.ҽ��ID(+) " & _
      "             And A.����=D.����(+)" & str��ʾ���� & _
      " Order by A.����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        Val(mArrFilter("����ID")), _
        Val(mArrFilter("ҽ��ID")(0)), _
        CStr(mArrFilter("ҽ��ID")(0)), _
        CDate(mArrFilter("��Ч��")(0)), CDate(mArrFilter("��Ч��")(1)), UserInfo.ID)
    mblnNotMoveRow = True
    If Not mrsRoom Is Nothing Then
        If mrsRoom.State = 1 Then mrsRoom.Close
    End If
    Set mrsRoom = Nothing
    
    With Me.vsList
        If .Row > 0 Then
            lngPreID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End If
        .Clear 1
        .Rows = 3: lngRow = 2
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("״̬")) = IIf(Val(Nvl(rsTemp!�Ƿ�ɾ��, 0)) = 1, "��ɾ", IIf(Nvl(rsTemp!ͣ������) <> "", "��ͣ", "����"))
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTemp!�ű�)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(rsTemp!��Ŀ)
            .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ��)
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            If IsNull(rsTemp!������Լ) Then
                .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            Else
                If Val(rsTemp!������Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("��һ-����")) = Nvl(rsTemp!��һ)
            .TextMatrix(lngRow, .ColIndex("��һ-�޺�")) = Format(Val(Nvl(rsTemp!��һ�޺�)), "###;;")
            If IsNull(rsTemp!��һ��Լ) Then
                .TextMatrix(lngRow, .ColIndex("��һ-��Լ")) = Format(Val(Nvl(rsTemp!��һ��Լ)), "###;;")
            Else
                If Val(rsTemp!��һ��Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("��һ-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("��һ-��Լ")) = Format(Val(Nvl(rsTemp!��һ��Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("�ܶ�-����")) = Nvl(rsTemp!�ܶ�)
            .TextMatrix(lngRow, .ColIndex("�ܶ�-�޺�")) = Format(Val(Nvl(rsTemp!�ܶ��޺�)), "###;;")
            If IsNull(rsTemp!�ܶ���Լ) Then
                .TextMatrix(lngRow, .ColIndex("�ܶ�-��Լ")) = Format(Val(Nvl(rsTemp!�ܶ���Լ)), "###;;")
            Else
                If Val(rsTemp!�ܶ���Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("�ܶ�-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("�ܶ�-��Լ")) = Format(Val(Nvl(rsTemp!�ܶ���Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            If IsNull(rsTemp!������Լ) Then
                .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            Else
                If Val(rsTemp!������Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            If IsNull(rsTemp!������Լ) Then
                .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            Else
                If Val(rsTemp!������Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            If IsNull(rsTemp!������Լ) Then
                .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            Else
                If Val(rsTemp!������Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            If IsNull(rsTemp!������Լ) Then
                .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            Else
                If Val(rsTemp!������Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("����")) = IIf(Val(Nvl(rsTemp!����)) = 0, "", "��")
            .TextMatrix(lngRow, .ColIndex("���﷽ʽ")) = Nvl(rsTemp!���﷽ʽ)
            .TextMatrix(lngRow, .ColIndex("�ѹ���")) = Nvl(rsTemp!�ѹ���)
            .TextMatrix(lngRow, .ColIndex("��Լ��")) = Nvl(rsTemp!��Լ��)
            '59611 ���ϴ� 2014/04/11 16:40:49 ���һ����ʾԤԼ�ѽ�����
            .TextMatrix(lngRow, .ColIndex("�����ѽ�����")) = Nvl(rsTemp!�����ѽ���)
             
            .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!����ID) & "_" & Nvl(rsTemp!��ĿID) & "_" & Nvl(rsTemp!ҽ��ID)
            .TextMatrix(lngRow, .ColIndex("Ӧ������")) = Read����Ӧ������(Val(Nvl(rsTemp!ID)))    ' Nvl(rsTemp!��������)
            If Not IsNull(rsTemp!��ʼʱ��) Then
                .TextMatrix(lngRow, .ColIndex("��Ч��Χ")) = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & _
                    "��" & Format(rsTemp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(lngRow, .ColIndex("��Ч��Χ")) = Replace(.TextMatrix(lngRow, .ColIndex("��Ч��Χ")), " 00:00:00", "")
            End If
            .TextMatrix(lngRow, .ColIndex("��ſ���")) = IIf(Val(Nvl(rsTemp!��ſ���)) = 0, "", "��")
            .TextMatrix(lngRow, .ColIndex("ͣ������")) = Nvl(rsTemp!ͣ������)
            If Trim(.TextMatrix(lngRow, .ColIndex("ͣ������"))) <> "" Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
            End If
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "��Ч��-�ű��б�", True
        .ColWidth(.ColIndex("��־")) = 285
        If lngPreID <> 0 Then
            lngRow = .FindRow(lngPreID, 0, .ColIndex("ID"), , True)
            If lngRow > 0 And lngRow <= .Rows - 1 Then .Row = lngRow
        End If
        If .Row <= 0 Then
            .Row = 1
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        mblnNotMoveRow = False
        '��ȡ��ϸ:
        Call vsList_AfterRowColChange(0, 0, .Row, .Col)
    End With
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsList.Redraw = flexRDBuffered
End Sub

Public Property Get zlHaveDatePlan(Optional blnPlanID As Boolean = False) As Boolean
    If Not blnPlanID Then
         zlHaveDatePlan = mblnHaveDate
         Exit Property
    End If
     zlHaveDatePlan = mfrm�ƻ�.zlHaveDatPlan
End Property

 

Private Function Read����Ӧ������(ByVal lngID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������
    '���:lngID-ID
    '����:
    '����:
    '����:���˺�
    '����:2009-09-14 22:39:14
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo errH
    If lngID = 0 Then Exit Function
    
    If mrsRoom Is Nothing Then
        strSQL = "Select ��������,�ű�ID From �ҺŰ�������"
        Set mrsRoom = New Recordset
        Call zlDatabase.OpenRecordset(mrsRoom, strSQL, Me.Caption)
    End If
    
    With mrsRoom
        .Filter = "�ű�ID=" & lngID
        If .RecordCount = 0 Then Exit Function
        
        Do While Not .EOF
            Read����Ӧ������ = Read����Ӧ������ & ";" & !��������
            .MoveNext
        Loop
    End With
    Read����Ӧ������ = Mid(Read����Ӧ������, 2)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Read�ƻ�Ӧ������(ByVal lng����ID As Long, ByVal lng�ƻ�ID As Long, Optional blnReRead As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������
    '���:lngID-ID
    '     blnReRead-���¶�ȡ
    '����:
    '����:
    '����:���˺�
    '����:2009-09-14 22:39:14
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo errH
    If lng����ID = 0 Then Exit Function
    
    If mrsPlanRoom Is Nothing Or blnReRead Then
        strSQL = "Select ��������,�ƻ�ID From �Һżƻ����� A,�ҺŰ��żƻ� B where a.�ƻ�id=B.ID and b.����ID=[1]"
        Set mrsPlanRoom = New Recordset
        Set mrsPlanRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    End If
    
    With mrsPlanRoom
        .Filter = "�ƻ�ID=" & lng�ƻ�ID
        If .RecordCount = 0 Then Exit Function
        
        Do While Not .EOF
            Read�ƻ�Ӧ������ = Read�ƻ�Ӧ������ & ";" & Nvl(!��������)
            .MoveNext
        Loop
    End With
    Read�ƻ�Ӧ������ = Mid(Read�ƻ�Ӧ������, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Property Get GetList() As VSFlexGrid
    Set GetList = vsList
End Property
Public Function zlExecuteDeleteList(ByVal blnԤԼ����ֹɾ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ָ����Ч�ű�����
    '���:blnԤԼ����ֹɾ��-��ԤԼ��,��ֹɾ��
    '����:ɾ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-15 10:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, str�ű� As String
    Dim intOper As Integer
    With vsList
        Err = 0: On Error GoTo Errhand
        If MsgBox("��ȷ��Ҫɾ���ű�""" & .TextMatrix(.Row, .ColIndex("�ű�")) & """�İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            
            intRow = .Row
            str�ű� = .TextMatrix(.Row, .ColIndex("�ű�"))
            If CheckExistsBooking(str�ű�) Then
                '����:46639
                If blnԤԼ����ֹɾ�� Then
                    Call MsgBox("�úű����ԤԼ�Һŵ�,����ɾ��!", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
                If MsgBox("�úű����ԤԼ�Һŵ�,��ȷʵҪɾ����?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            '�޸�������
            intOper = CheckExistsRegster(str�ű�)
            If intOper = 0 Then Exit Function '����Ƿ�ð������йҺ�����
            If intOper = 1 Then '1Ϊ��ɾ��
                strSQL = "zl_�ҺŰ���_Delete(" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & ",0)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            ElseIf intOper = 2 Then '2Ϊ��ɾ��
                strSQL = "zl_�ҺŰ���_Delete(" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & ",1)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            If mArrFilter("��ʾɾ������") And intOper = 1 Then
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
                .Cell(flexcpText, .Row, .ColIndex("״̬")) = "��ɾ"
                .Cell(flexcpText, .Row, .ColIndex("ͣ������")) = Format(Date, "yyyy-mm-dd")
                Exit Function
            End If
            If .Rows > 2 Then
                .RemoveItem intRow
            Else
                For i = 0 To .Cols - 1
                    .TextMatrix(intRow, i) = ""
                Next
            End If
            If intRow <= .Rows - 1 Then
                .Row = intRow
            Else
                .Row = .Rows - 1
            End If
            .Col = 0: .ColSel = .Cols - 1
        End If
    End With
    If mblnListSetFocus Then
        zlControl.ControlSetFocus vsList
    Else
        mfrm�ƻ�.SetGotFocus (True)
    End If
    zlExecuteDeleteList = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlExecuteModifyList(ByVal frmMain As Form) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸�ָ����Ч�ű�����
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-15 10:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�ű� As String, intRow As Integer, lngID As Long
    With vsList
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Function
        If InStr(1, mstrPrivs, ";����;") < 0 Then Exit Function
        
        intRow = .Row
        str�ű� = .TextMatrix(.Row, .ColIndex("�ű�"))
        If CheckExistsBooking(str�ű�) Then
           If MsgBox("�úű����ԤԼ�Һŵ�,��ȷ��Ҫ�޸���", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Function
            End If
        End If
        If frmRegistPlanEdit.ShowEdit(frmMain, edt_�޸�, mlngModule, mstrPrivs, lngID) = False Then
            Exit Function
        End If
        Call LoadDataToList
        zlControl.ControlSetFocus vsList, True
        zlExecuteModifyList = True
    End With
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckExistsBooking(str�ű� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ű��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Min(����ʱ��) ʱ��" & vbNewLine & _
            "From ������ü�¼" & vbNewLine & _
            "Where ��¼���� = 4 And ��¼״̬ In (0, 1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
'    If gintԤԼ���� = 0 Then
    strSQL = strSQL & " And ����ʱ�� > Sysdate"
'    Else
'        strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
'    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�)
    
    CheckExistsBooking = Not IsNull(rsTmp!ʱ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Sub zlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ص��Զ��屨��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-15 11:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, str���� As String
    '����ID_��ĿID_ҽ��ID
    With vsList
        varData = Split(.TextMatrix(.Row, .ColIndex("IDS")) & "___", "_")
        str���� = Trim(.TextMatrix(.Row, .ColIndex("����")))
        If str���� <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain, _
                "����=" & str����, "�ű�=" & Trim(.TextMatrix(.Row, .ColIndex("�ű�"))), _
                "����=" & Val(varData(0)), _
                "��Ŀ=" & Val(varData(1)), _
                "ҽ��=" & Val(varData(2)))
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
        End If
    End With
End Sub
Public Property Get zlGetListCurrRow() As Variant
    Dim varData As Variant
    Dim cllRow As New Collection
    
    '����ID_��ĿID_ҽ��ID
    With vsList
        varData = Split(.TextMatrix(.Row, .ColIndex("IDS")) & "___", "_")
        cllRow.Add Val(varData(0)), "����ID"
        cllRow.Add Val(varData(1)), "��ĿID"
        cllRow.Add Val(varData(2)), "ҽ��ID"
        cllRow.Add Trim(.TextMatrix(.Row, .ColIndex("����"))), "����"
        cllRow.Add Trim(.TextMatrix(.Row, .ColIndex("�ű�"))), "�ű�"
    End With
    Set zlGetListCurrRow = cllRow
End Property
Public Property Get zlGet����ID(Optional blnPlanID As Boolean = False) As Long
    If blnPlanID Then
        zlGet����ID = mfrm�ƻ�.zlGet�ƻ�ID: Exit Sub
    End If
    With vsList
        zlGet����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property

Public Property Get zlPlanStatus() As Long
   zlPlanStatus = mfrm�ƻ�.zlPlanStatus
    '��ȡ�ƻ����ŵĵ�ǰ״̬
    '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч
End Property

Public Sub zlRefreshData(ByVal ArrFilter As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mArrFilter = ArrFilter
    Call LoadDataToList
    Call zlActtion
End Sub

Public Sub zlActtion()
    If mblnListSetFocus = True Then
        On Error Resume Next
        If vsList.Visible And vsList.Enabled Then vsList.SetFocus
    Else
        mfrm�ƻ�.zlActtion
    End If
End Sub

Private Sub vsPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent zlPopuMenu(1, Button, Shift, X, Y)
End Sub
Public Sub zlRefreshOlnyPlanData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ˢ�¼ƻ�����
    '����:���˺�
    '����:2009-09-17 11:28:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, str���� As String, str�ű� As String, str���� As String
    With vsList
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        str���� = Trim(.TextMatrix(.Row, .ColIndex("����")))
        str�ű� = Trim(.TextMatrix(.Row, .ColIndex("�ű�")))
        str���� = Trim(.TextMatrix(.Row, .ColIndex("����")))
    End With
    Call mfrm�ƻ�.zlShowPlan(lng����ID, str����, str�ű�, str����)
    Call LoadStopPlan(lng����ID)
End Sub

Private Function CheckIsUserPreRegist(ByVal str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����ԤԼ�Һ�
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2011-02-24 09:53:25
    '����:35959
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 1 From ������ü�¼ where ����ʱ��>=sysdate And ��¼״̬=0 and  ���㵥λ=[1]  and  ��¼����=4 and rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If

    strSQL = "Select 1 From ���˹Һż�¼ where ����ʱ��>=sysdate and ��¼����=1 and ��¼״̬=1 and  �ű�=[1]  and rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If
    CheckIsUserPreRegist = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Public Function zlStopAndResume(ByVal blnStop As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ͣ�û�����ָ����Ч�ű�����
    '���أ�ͣ�û����óɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-08-17 11:19:18
    '˵����31923
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, str�ű� As String
    With vsList
        Err = 0: On Error GoTo Errhand
        If blnStop Then
            '����:35959
            If CheckIsUserPreRegist(.TextMatrix(.Row, .ColIndex("�ű�"))) Then
                If MsgBox("ע��:" & vbCrLf & "   �ű�Ϊ" & .TextMatrix(.Row, .ColIndex("�ű�")) & "�İ����Ѿ�����ԤԼ,�Ƿ����ͣ��? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        End If
        
        If MsgBox("��ȷ��Ҫ" & IIf(blnStop, "ͣ��", "����") & "�ű�""" & .TextMatrix(.Row, .ColIndex("�ű�")) & """�İ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            intRow = .Row
            str�ű� = .TextMatrix(.Row, .ColIndex("�ű�"))
            strSQL = "zl_�ҺŰ���_StopAndStart(" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "," & IIf(blnStop, 1, 0) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            If blnStop Then
                .TextMatrix(.Row, .ColIndex("ͣ������")) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
            Else
                .TextMatrix(.Row, .ColIndex("ͣ������")) = ""
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = .ForeColor
            End If
            If mArrFilter("��ʾͣ�ð���") = 1 Or blnStop = False Then
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = IIf(blnStop, vbRed, vbBlack)
                .Cell(flexcpText, .Row, .ColIndex("״̬")) = IIf(blnStop = True, "��ͣ", "����")
                Exit Function
            End If
            If .Rows > 2 Then
                .RemoveItem intRow
            Else
                For i = 0 To .Cols - 1
                    .TextMatrix(intRow, i) = ""
                Next
            End If
            If intRow <= .Rows - 1 Then
                .Row = intRow
            Else
                .Row = .Rows - 1
            End If
            .Col = 0: .ColSel = .Cols - 1
        End If
    End With
    Call zlActtion
    zlStopAndResume = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlStopPlanTimes() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ͣ�ð��żƻ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-08 14:11:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, lngID As Long
    Err = 0: On Error GoTo Errhand
    With vsList
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Function
    End With
    If frmRegistPlanInvalidation.ShowCard(Me, mlngModule, mstrPrivs, lngID) = False Then
        Call zlActtion: Exit Function
    End If
    Call zlActtion
    zlStopPlanTimes = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlIsStopPlan() As Boolean
    '��ȡ�úű��Ƿ��Ѿ���ͣ��
    With vsList
        If .Row < 0 Then Exit Function
        zlIsStopPlan = Trim(.TextMatrix(.Row, .ColIndex("ͣ������"))) <> ""
    End With
End Function

 

Private Sub vsStop_GotFocus()
    vsStop.BackColorSel = &H8000000D
    mblnListSetFocus = True
End Sub

Private Sub vsStop_LostFocus()
    vsStop.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Public Function zlClearStopPlanTimes() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ͣ�üƻ�����
    '����:ɾ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-09 15:15:32
    '����:32504
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intMsg As VbMsgBoxResult
    Dim int��־ As Integer, lng����ID As Long
    
    '�������ͣ�üƻ�
    intMsg = MsgBox("���Ƿ����Ҫ������йҺŰ��ŵ�ͣ�üƻ�?" & vbCrLf & _
                    "���ǡ���ʾɾ�������ƶ��õ�ͣ�üƻ���" & vbCrLf & _
                    "���񡿱�ʾɾ�������Ѿ�ʧЧ�˵�ͣ�üƻ���" & vbCrLf & _
                    "��ȡ������ʾ��ɾ����" & vbCrLf & _
                    " ", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
    If intMsg = vbCancel Then Exit Function
    If intMsg = vbYes Then
        int��־ = 2
    Else
        int��־ = 1
    End If
    On Error GoTo errHandle
    'Zl_�ҺŰ���ͣ��״̬_Clearall
    strSQL = "Zl_�ҺŰ���ͣ��״̬_Clearall(" & int��־ & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If int��־ = 2 Then
        vsStop.Clear 1: vsStop.Rows = 2
    Else
        With vsList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
        Call LoadStopPlan(lng����ID)
    End If
    zlClearStopPlanTimes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2011-11-14 14:52:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    Set mfrm�ƻ� = New frmRegistPlanIntend
    
    Set ObjItem = tbPage.InsertItem(mPgIndex.pg_�ƻ�, "���żƻ���Ϣ", mfrm�ƻ�.Hwnd, 0)
    ObjItem.Tag = mPgIndex.pg_�ƻ�
    Set ObjItem = tbPage.InsertItem(mPgIndex.pg_ʱ��, "����ʱ����Ϣ", picTime.Hwnd, 0)
    ObjItem.Tag = mPgIndex.pg_ʱ��
    Set ObjItem = tbPage.InsertItem(mPgIndex.pg_ͣ��, "�ƻ�ͣ����Ϣ", picStop.Hwnd, 0)
    ObjItem.Tag = mPgIndex.pg_ͣ��
    If mblnHaveUnit Then
        Set ObjItem = tbPage.InsertItem(mPgIndex.pg_��λ, "������λ������Ϣ", frmUnitReg.Hwnd, 0)
        ObjItem.Tag = mPgIndex.pg_��λ
    End If
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zl_ReLoadUnitReg()
    Dim lng����ID As Long
    lng����ID = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ID")))
    If lng����ID <= 0 Then Exit Sub
    LoadUnitReg lng����ID
End Sub

Public Sub UpdatePara(ByVal blnPara As Boolean)
    mfrm�ƻ�.blnShowExpired = blnPara
End Sub

Public Property Get zlGet����ͣ��() As Boolean
    '��ȡ�����Ƿ��Ѿ�ͣ��
      If vsList.Row >= 0 And vsList.Col >= 0 Then
        zlGet����ͣ�� = vsList.Cell(flexcpForeColor, vsList.Row, vsList.Col) = vbRed
      End If
End Property

Private Function CheckExistsRegster(str�ű� As String) As Byte
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ҺŰ����Ƿ�����ѹҺ�����
    '���:lng����ID-����ID
    '����:0������,1��ɾ��,2��ɾ��
    '����:����
    '����:2012-03-12 10:10:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
     Dim rsTmp As ADODB.Recordset, strSQL As String, strRS As String, blnRs As Boolean
     strSQL = "Select Nvl(Max(A.�ѹҺ���),0) As �ѹҺ��� ,Nvl(Max(A.�ƻ���),0) As �ƻ��� From(" & vbNewLine & _
              "Select 1 As �ѹҺ���, 0 As �ƻ��� From ���˹ҺŻ��� A Where Rownum=1 And A.����=[1] Having Sum(Nvl(�ѹ���,0)) > 0 or Sum(Nvl(��Լ��,0)) > 0" & vbNewLine & _
              "Union All" & vbNewLine & _
              "Select 0 As �ѹҺ���, 1 As �ƻ��� From �ҺŰ��żƻ� C Where C.����=[1] Having Count(1)>1) A"
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�)
     If rsTmp!�ѹҺ��� > 0 And rsTmp!�ƻ��� > 0 Then
        '67824:������,2013-11-21,�Ի�����Ų���������
        If MsgBox("�ð������Ѿ��йҺ�������ƻ�����,��ȷ��Ҫɾ���ð���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        CheckExistsRegster = 1
        Exit Function
     End If
     If rsTmp!�ѹҺ��� > 0 Then
        If MsgBox("���Ű����Ѿ��йҺ�����,��ȷ��Ҫɾ���ð���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        CheckExistsRegster = 1
        Exit Function
     End If
     If rsTmp!�ƻ��� > 0 Then
        If MsgBox("���Ű����Ѿ��мƻ�����,��ȷ��Ҫɾ���ð���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        CheckExistsRegster = 2
        Exit Function
     End If
       If MsgBox("���ɾ�����Ű����ݽ��޷��ָ�,��ȷ��Ҫɾ���ð���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
       CheckExistsRegster = 2
       Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
