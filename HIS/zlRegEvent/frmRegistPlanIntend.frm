VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanIntend 
   BorderStyle     =   0  'None
   Caption         =   "���żƻ�"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTbPage 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5160
      ScaleHeight     =   2175
      ScaleWidth      =   3255
      TabIndex        =   6
      Top             =   1560
      Width           =   3255
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2295
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4575
         _Version        =   589884
         _ExtentX        =   8070
         _ExtentY        =   4048
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPlan 
      BorderStyle     =   0  'None
      Height          =   2520
      Left            =   780
      ScaleHeight     =   2520
      ScaleWidth      =   6645
      TabIndex        =   3
      Top             =   3090
      Width           =   6645
      Begin VSFlex8Ctl.VSFlexGrid vsPlan 
         Height          =   2145
         Left            =   -150
         TabIndex        =   4
         Top             =   30
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
         Cols            =   26
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanIntend.frx":0000
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   195
            Width           =   210
            Begin VB.Image imgColPlan 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistPlanIntend.frx":032F
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picTime 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   1425
      ScaleHeight     =   2745
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      Begin MSComctlLib.TabStrip tbWeekTime 
         Height          =   285
         Left            =   0
         TabIndex        =   1
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
         TabIndex        =   2
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
         FormatString    =   $"frmRegistPlanIntend.frx":087D
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   75
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRegistPlanIntend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnNotMoveRow As Boolean '���ƶ���
Private mrsRoom As ADODB.Recordset
Private mrsPlanRoom As ADODB.Recordset
Private Const conPane_Plan = 1
Private Const conPane_Time = 2
Private Const conPane_Unit = 3
'�����:51156
Private Const conTbPage = 4
Private mblnShowExpired As Boolean
Private mlng����ID As Long  '����ID
Private mstr���� As String, mstr�ű� As String, mstr���� As String
Public Event zlPopuMenu(intType As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PlanGotFocus(blnPlan As Boolean)
Public Event PlanLostFocus(blnPlan As Boolean)
Private Enum mPgIndex
    pg_�ƻ� = 1
    pg_ʱ�� = 2
    pg_ͣ�� = 3
End Enum
Private mblnNotClick As Boolean
Private mrsTime As ADODB.Recordset
Private mblnHaveDatPlan As Boolean
Private mfrmUnitReg As frmUnitRegPlan
Public mblnSelected�ƻ� As Boolean '�����:51429

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
      
    Set panThis = dkpMan.CreatePane(conPane_Plan, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "�ƻ���Ϣ"
    panThis.Handle = picPlan.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = conPane_Plan
     '�����:51156
    Set panThis = dkpMan.CreatePane(conTbPage, 250, 580, DockRightOf, panThis)
    panThis.Title = "�ƻ�ʱ����Ϣ�������λ�ƻ�������Ϣ"
    panThis.Tag = conTbPage
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption
    panThis.Handle = picTbPage.Hwnd
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    Call InitTbPage
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub
'�����:51156
Private Sub InitTbPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2011-11-14 14:52:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    If mfrmUnitReg Is Nothing Then Set mfrmUnitReg = New frmUnitRegPlan
    
    Set ObjItem = tbPage.InsertItem(conPane_Time, "�ƻ�ʱ����Ϣ", picTime.Hwnd, 0)
    ObjItem.Tag = conPane_Time
    Set ObjItem = tbPage.InsertItem(conPane_Unit, "������λ�ƻ�������Ϣ", mfrmUnitReg.Hwnd, 0)
    ObjItem.Tag = conPane_Unit
     With tbPage
         tbPage.Item(0).Selected = True
        .PaintManager.Position = xtpTabPositionBottom
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
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocked Or Action = PaneActionDocking Then Exit Sub
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Plan
        Item.Handle = picPlan.Hwnd
    Case conPane_Time
        Item.Handle = picTime.Hwnd
    Case conTbPage '�����:51156
        Item.Handle = picTbPage.Hwnd
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
    With vsPlan
        .Redraw = flexRDNone
        .Rows = 3: .FixedRows = 2
        .FixedCols = 1
        .Cols = 40:   .Clear
        .FrozenCols = 2
        .TextMatrix(0, i) = "  ": .ColWidth(i) = 285
        .TextMatrix(1, i) = "  ":  .ColKey(i) = "��־": i = i + 1
        
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0
        .TextMatrix(1, i) = "ID": .ColKey(i) = "ID": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 720: .ColHidden(i) = True:
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1

        .TextMatrix(0, i) = "�ű�": .ColWidth(i) = 480: .ColHidden(i) = True:
        .TextMatrix(1, i) = "�ű�": .ColKey(i) = "�ű�": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000: .ColHidden(i) = True:
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1
        .TextMatrix(0, i) = "��Ŀ": .ColWidth(i) = 1000: .ColHidden(i) = True:
        .TextMatrix(1, i) = "��Ŀ": .ColKey(i) = "��Ŀ": i = i + 1
        .TextMatrix(0, i) = "ҽ��":: .ColWidth(i) = 1000
        .TextMatrix(1, i) = "ҽ��": .ColKey(i) = "ҽ��": i = i + 1
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
        .TextMatrix(0, i) = "��Чʱ��": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "��Чʱ��": .ColKey(i) = "��Чʱ��": i = i + 1
        .TextMatrix(0, i) = "ʧЧʱ��": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "ʧЧʱ��": .ColKey(i) = "ʧЧʱ��": i = i + 1
        .TextMatrix(0, i) = "��ſ���": .ColWidth(i) = 765
        .TextMatrix(1, i) = "��ſ���": .ColKey(i) = "��ſ���": i = i + 1
        
        .TextMatrix(0, i) = "������": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "������": .ColKey(i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "����ʱ��": .ColKey(i) = "����ʱ��": i = i + 1
        
        .TextMatrix(0, i) = "�����": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "�����": .ColKey(i) = "�����": i = i + 1
        .TextMatrix(0, i) = "���ʱ��": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "���ʱ��": .ColKey(i) = "���ʱ��": i = i + 1
        .TextMatrix(0, i) = "ʵ��ִ��ʱ��": .ColWidth(i) = 1500
        .TextMatrix(1, i) = "ʵ��ִ��ʱ��": .ColKey(i) = "ʵ��ִ��ʱ��": i = i + 1
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
            Case "����", "�ű�", "��Чʱ��"
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
    Call InitVsGrid
    Call vsPlan_LostFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "��Ч��-�ƻ��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
 

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsPlan, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "��Ч��-�ƻ��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
    If Not mfrmUnitReg Is Nothing Then Unload mfrmUnitReg: Set mfrmUnitReg = Nothing
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
'�����:51156
Private Sub picTbPage_Resize()
    tbPage.Move picTbPage.Left, picTbPage.Top, picTbPage.Width, picTbPage.Height
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
        vsTime.Height = .ScaleHeight - vsTime.Top
    End With
End Sub
Private Sub picPlan_Resize()
    Err = 0: On Error Resume Next
    With picPlan
        vsPlan.Left = .ScaleLeft
        vsPlan.Width = .ScaleWidth
        vsPlan.Top = .ScaleTop
        vsPlan.Height = .ScaleHeight
    End With
End Sub

Public Function zlHaveDatPlan() As Boolean
    '*************************************
    'ʱ���Ƿ�������
    '*************************************
   
    zlHaveDatPlan = mblnHaveDatPlan
End Function

Public Function zlHaveData() As Boolean
    '����:�Ƿ�������
    With Me.vsPlan
        zlHaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
    End With
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
    If mstr�ű� = "" Or mlng����ID = 0 Then Exit Sub
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "�Һżƻ���"
    objRow.Add "���ࣺ" & mstr����
    objRow.Add "�ű�" & mstr�ű�
    objRow.Add "���ң�" & mstr����
 
    objPrint.UnderAppRows.Add objRow
    Set vsGrid = vsPlan
        
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
Public Function zlShowPlan(ByVal lng����ID As Long, ByVal str���� As String, ByVal str�ű� As String, ByVal str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�ƻ�������Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-11-15 13:54:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstr���� = str����: mstr�ű� = str�ű�: mstr���� = str����: mlng����ID = lng����ID
    Call LoadPlan(lng����ID)
    zlShowPlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Sub ReloadTimePlan()
    '***********************************************
    ' ���¼��ؼƻ���ʱ���
    '***********************************************
    Dim lng�ƻ�ID As Long, bln��ſ��� As Boolean
    With vsPlan
        lng�ƻ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln��ſ��� = Trim(.TextMatrix(.Row, .ColIndex("��ſ���"))) <> ""
    End With
    Call LoadTimePlan(lng�ƻ�ID, bln��ſ���, True)
End Sub
Private Sub LoadTimePlan(ByVal lng�ƻ�ID As Long, ByVal bln��ſ��� As Boolean, Optional blnReload As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͣ�üƻ�����
    '����:���˺�
    '����:2010-09-09 11:54:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str���� As String
    Dim i As Long, r As Integer, strʱ�� As String, strTime As String, strKey As String
    Static lngPre�ƻ�Id  As Long
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
        lngPre�ƻ�Id = -1
    ElseIf mrsTime.State <> 1 Then
         lngPre�ƻ�Id = -1
    End If
    If lngPre�ƻ�Id <> lng�ƻ�ID Or blnReload Then
        lngPre�ƻ�Id = lng�ƻ�ID
        strSQL = "" & _
        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               ��������,�Ƿ�ԤԼ" & _
        "   From  �Һżƻ�ʱ�� " & _
        "   Where �ƻ�ID=[1] And ��������>0 " & _
        "   Order by ����,ʱ��,���"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ƻ�ID)
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
             r = 0: i = 0
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
 
Private Sub tbWeekTime_Click()
   Dim lng�ƻ�ID As Long, bln��ſ��� As Boolean
    With vsPlan
        lng�ƻ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln��ſ��� = Trim(.TextMatrix(.Row, .ColIndex("��ſ���"))) <> ""
    End With
    Call LoadTimePlan(lng�ƻ�ID, bln��ſ���)
End Sub

Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng�ƻ�ID As Long, bln��ſ��� As Boolean
    If OldRow = NewRow Then Exit Sub
    With vsPlan
        lng�ƻ�ID = Val(.TextMatrix(NewRow, .ColIndex("ID")))
        bln��ſ��� = Trim(.TextMatrix(NewRow, .ColIndex("��ſ���"))) <> ""
    End With
    Call LoadTimePlan(lng�ƻ�ID, bln��ſ���)
    Call LoadUnitPlan(lng�ƻ�ID)
    mblnSelected�ƻ� = True '�����:51429
End Sub

Private Sub LoadUnitPlan(ByVal lng�ƻ�ID As Long)
    If mfrmUnitReg Is Nothing Then Exit Sub
    mfrmUnitReg.ShowUnitReg lng�ƻ�ID
    
End Sub
Public Sub ReLoadUnitPlan()
    Dim lng�ƻ�ID As Long
    
    If mfrmUnitReg Is Nothing Then Exit Sub
    With vsPlan
        lng�ƻ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
    mfrmUnitReg.ShowUnitReg lng�ƻ�ID
    
End Sub
Private Sub vsPlan_BeforeMoveColumn(ByVal Col As Long, Position As Long)
        With vsPlan
            If .ColKey(Col) Like "��*" Then
                 Position = Col
            End If
        End With
End Sub
Private Sub vsPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "��Ч��-�ƻ��б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub
Public Sub SetGotFocus(blnPlan As Boolean)
    If blnPlan Then vsPlan_GotFocus: Exit Sub
    vsTime_GotFocus
End Sub
Public Sub SetLostFocus(blnPlan As Boolean)
    If blnPlan Then vsPlan_LostFocus: Exit Sub
    vsTime_LostFocus
End Sub
Private Sub vsPlan_GotFocus()
    vsPlan.BackColorSel = &H8000000D
    RaiseEvent PlanGotFocus(True)
End Sub

Private Sub vsPlan_LostFocus()
    vsPlan.BackColorSel = GRD_LOSTFOCUS_COLORSEL
    RaiseEvent PlanLostFocus(True)
End Sub
Private Sub vsTime_GotFocus()
    RaiseEvent PlanGotFocus(False)
End Sub
Private Sub vsTime_LostFocus()
    RaiseEvent PlanLostFocus(False)
End Sub
Private Sub LoadPlan(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ذ��żƻ�
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngRow As Long
    Dim strSQL As String, lngPreID As Long, strTable As String
    Dim lng���� As Long
    If lng����ID = 0 Then
        vsPlan.Clear 1
        vsPlan.Rows = 2: Exit Sub
    End If
    
    Err = 0: On Error GoTo Errhand:
    strTable = "" & _
    "   Select A.ID, " & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'��һ',B.�޺���,0)) as ��һ�޺�, Sum(Decode(B.������Ŀ,'��һ',B.��Լ��))  as ��һ��Լ," & _
    "             Sum(Decode(B.������Ŀ,'�ܶ�',B.�޺���,0)) as �ܶ��޺�, Sum(Decode(B.������Ŀ,'�ܶ�',B.��Լ��))  as �ܶ���Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��))  as ������Լ" & _
    "   From �ҺŰ��żƻ� A,�Һżƻ����� B  " & _
    "   Where A.ID=B.�ƻ�ID(+)  and A.����ID=[1] And A.��Чʱ�� >= A.����ʱ��-3/24/60/60" & _
    IIf(mblnShowExpired, "", " And A.ʧЧʱ�� > Sysdate") & _
    "   Group by A.ID"
    '֮���԰���ʱ���ȥ����Ϊ�˱���,������Ч�������
   '38505
    strSQL = " " & _
        "   Select P.*,B.���� As ��Ŀ,D.���� As ���� " & _
        "   From ( " & _
        "     Select  row_number()  over (Partition By �ƻ�id Order By �ƻ�id,���� Desc) As ���1,M.* " & _
        "     From ( " & _
        "       Select Level As ����, Sys_Connect_By_Path(��������, ';') �������Ҽ�, Q.*  " & _
        "       From ( Select  C.Id as �ƻ�ID,C.����ID ,A.����,  A.����,  A.����id,   Nvl(C.��Ŀid,a.��ĿID) as ��ĿID, C.ҽ������,  C.ҽ��id,   " & _
        "                              C.����,C1.�����޺�,C1.������Լ,C.��һ,C1.��һ�޺�,C1.��һ��Լ,C.�ܶ�,C1.�ܶ��޺�,C1.�ܶ���Լ, " & _
        "                              C.����,C1.�����޺�,C1.������Լ,C.����,C1.�����޺�,C1.������Լ,C.����,C1.�����޺�,C1.������Լ, " & _
        "                              C.����,C1.�����޺�,C1.������Լ, " & _
        "                              A.��������,   Decode(Nvl(C.���﷽ʽ,0),0,'������',1,'ָ������',2,'��̬����',3,'ƽ������') as ���﷽ʽ ,  C.��ſ���,             " & _
        "                              to_char(A.��ʼʱ��,'yyyy-mm-dd hh24:mi:ss') ��ʼʱ��,  to_char(A.��ֹʱ��,'yyyy-mm-dd hh24:mi:ss') ��ֹʱ��,       " & _
        "                              to_char(C.��Чʱ��,'yyyy-mm-dd hh24:mi:ss') as ��Чʱ��,to_char(C.ʧЧʱ��,'yyyy-mm-dd hh24:mi:ss') as ʧЧʱ��,            " & _
        "                              to_char(C.ʵ����Ч,'yyyy-mm-dd hh24:mi:ss') as ʵ��ִ��ʱ��, C.������,to_char(C.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,            " & _
        "                              C.�����,to_char(C.���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ�� , " & _
        "                               b.��������,row_number() over (Partition By �ƻ�ID Order By �ƻ�id,��������) As ��� " & _
        "           From (" & strTable & ") C1,�ҺŰ��żƻ� C,�ҺŰ��� A,�Һżƻ����� B " & _
        "           Where  C.ID=C1.ID And C.����ID =A.Id And C.Id=B.�ƻ�ID(+) " & _
        "           Order By �ƻ�ID,�������� ) Q " & _
        "        Connect By �ƻ�id= Prior �ƻ�id And ���-1 =Prior ��� " & _
        "        )  M ) P,�շ���ĿĿ¼ B,���ű� D " & _
        "    Where P.���1=1 And P.��Ŀid=b.Id And P.����id =d.Id(+)  " & _
        "    Order By   ��Чʱ�� Desc, �ƻ�ID DESC "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    With Me.vsPlan
        If .Row > 0 And .Row <= .Rows - 1 Then
            lngPreID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End If
        .Clear 1
        .Rows = 3: .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 2
        lngRow = 2
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!�ƻ�Id)
            .Cell(flexcpData, lngRow, .ColIndex("ID")) = Nvl(rsTemp!����ID)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(rsTemp!��Ŀ)
            .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Nvl(rsTemp!�����޺�)
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
            .TextMatrix(lngRow, .ColIndex("��һ-�޺�")) = Nvl(rsTemp!��һ�޺�)
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
            .TextMatrix(lngRow, .ColIndex("�ܶ�-�޺�")) = Nvl(rsTemp!�ܶ��޺�)
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
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Nvl(rsTemp!�����޺�)
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
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Nvl(rsTemp!�����޺�)
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
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Nvl(rsTemp!�����޺�)
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
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Nvl(rsTemp!�����޺�)
            If IsNull(rsTemp!������Լ) Then
                .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            Else
                If Val(rsTemp!������Լ) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("����")) = IIf(Val(Nvl(rsTemp!��������)) = 0, "", "��")
            .TextMatrix(lngRow, .ColIndex("���﷽ʽ")) = Nvl(rsTemp!���﷽ʽ)
            .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!����ID) & "_" & Nvl(rsTemp!��ĿID) & "_" & Nvl(rsTemp!ҽ��ID)
             
            '*************************************************
            lng���� = Nvl(Nvl(rsTemp!�����޺�, 0)) + lng����
            lng���� = Nvl(Nvl(rsTemp!��һ�޺�, 0)) + lng����
            lng���� = Nvl(Nvl(rsTemp!�ܶ��޺�, 0)) + lng����
            lng���� = Nvl(Nvl(rsTemp!�����޺�, 0)) + lng����
            lng���� = Nvl(Nvl(rsTemp!�����޺�, 0)) + lng����
            lng���� = Nvl(Nvl(rsTemp!�����޺�, 0)) + lng����
            lng���� = Nvl(Nvl(rsTemp!�����޺�, 0)) + lng����
            '*************************************************
            If Nvl(rsTemp!�������Ҽ�) <> "" Then
                .TextMatrix(lngRow, .ColIndex("Ӧ������")) = Mid(Nvl(rsTemp!�������Ҽ�), 2)  'Read�ƻ�Ӧ������(lng����ID, Val(Nvl(rsTemp!�ƻ�ID)), False) ' Nvl(rsTemp!��������)
            End If
            
            If Not IsNull(rsTemp!��Чʱ��) Then
                .TextMatrix(lngRow, .ColIndex("��Чʱ��")) = Format(rsTemp!��Чʱ��, "yyyy-MM-dd HH:mm:ss")
                If Format(Nvl(rsTemp!��Чʱ��), "yyyy-MM-dd HH:mm:ss") <= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And Nvl(rsTemp!���ʱ��) <> "" Then
                    '�Ѿ���Ч,���ܸ���
                    .Cell(flexcpData, lngRow, .ColIndex("��Чʱ��")) = 1
                Else
                    'δ��Ч,�ܸ���
                    .Cell(flexcpData, lngRow, .ColIndex("��Чʱ��")) = 0
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("ʧЧʱ��")) = Nvl(rsTemp!ʧЧʱ��)
            .TextMatrix(lngRow, .ColIndex("��ſ���")) = IIf(Val(Nvl(rsTemp!��ſ���)) = 0, "", "��")
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("���ʱ��")) = Nvl(rsTemp!���ʱ��)
            If Nvl(rsTemp!ʵ��ִ��ʱ��) < "3000-01-01" Then
                .TextMatrix(lngRow, .ColIndex("ʵ��ִ��ʱ��")) = Nvl(rsTemp!ʵ��ִ��ʱ��)
            End If
            If Val(.Cell(flexcpData, lngRow, .ColIndex("��Чʱ��"))) = 1 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000010
            Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
            End If
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        If lngPreID <> 0 Then
            lngRow = .FindRow(lngPreID, 0, .ColIndex("ID"), , True)
            If lngRow > 0 Then .Row = lngRow
        Else
            .Row = 1
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        '�ָ�������
         zl_vsGrid_Para_Restore mlngModule, vsPlan, Me.Caption, "��Ч��-�ƻ��б�", True
        .ColWidth(.ColIndex("��־")) = 285
        Call vsPlan_AfterRowColChange(0, 0, .Row, .Col)
        .Redraw = flexRDBuffered
        mblnHaveDatPlan = lng���� > 0
    End With
   Exit Sub
Errhand:
    vsPlan.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
 
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
Public Property Get zlGet�ƻ�ID() As Long
    With vsPlan
        If .Row < 0 Or .Col < 0 Then Exit Sub
        zlGet�ƻ�ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property
Public Property Get zlPlanStatus() As Long
    Dim lngID As Long
    '��ȡ�ƻ����ŵĵ�ǰ״̬
    '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч
    With vsPlan
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then zlPlanStatus = 0: Exit Property
        If .TextMatrix(.Row, .ColIndex("���ʱ��")) <> "" Then
            zlPlanStatus = 2
            If Val(.Cell(flexcpData, .Row, .ColIndex("��Чʱ��"))) = 1 Then
                zlPlanStatus = 3
            End If
        Else
              zlPlanStatus = 1
        End If
    End With
End Property
Public Sub zlActtion()
    zlControl.ControlSetFocus vsPlan, True
End Sub
Private Sub vsPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent zlPopuMenu(1, Button, Shift, X, Y)
End Sub

Public Property Let blnShowExpired(ByVal vNewValue As Boolean)
    mblnShowExpired = vNewValue
End Property

