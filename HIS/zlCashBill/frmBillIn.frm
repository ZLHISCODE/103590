VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBillIn 
   Caption         =   "Ʊ��������"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9210
   Icon            =   "frmBillIn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPage 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   3600
      ScaleHeight     =   1710
      ScaleWidth      =   3225
      TabIndex        =   12
      Top             =   2940
      Width           =   3225
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2850
         Left            =   -1050
         TabIndex        =   13
         Top             =   315
         Width           =   3615
         _Version        =   589884
         _ExtentX        =   6376
         _ExtentY        =   5027
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDrawBill 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   5880
      ScaleHeight     =   2250
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   4155
      Width           =   3195
      Begin VSFlex8Ctl.VSFlexGrid vsDrawBill 
         Height          =   2145
         Left            =   0
         TabIndex        =   10
         Top             =   420
         Width           =   2895
         _cx             =   5106
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
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillIn.frx":0442
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
         Begin VB.PictureBox picColDrawBill 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   11
            Top             =   45
            Width           =   210
            Begin VB.Image imgColDrawBill 
               Height          =   195
               Left            =   0
               Picture         =   "frmBillIn.frx":05F2
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picDamnifyBill 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   2025
      ScaleHeight     =   2250
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   4740
      Width           =   3195
      Begin VSFlex8Ctl.VSFlexGrid vsDamnifyBill 
         Height          =   2145
         Left            =   0
         TabIndex        =   7
         Top             =   255
         Width           =   2895
         _cx             =   5106
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
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillIn.frx":0B40
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
         Begin VB.PictureBox picColDamnifyBill 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   8
            Top             =   45
            Width           =   210
            Begin VB.Image imgColDamnifyBill 
               Height          =   195
               Left            =   0
               Picture         =   "frmBillIn.frx":0C44
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picBillIn 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   1665
      ScaleHeight     =   2250
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   240
      Width           =   6795
      Begin VSFlex8Ctl.VSFlexGrid vsBillIn 
         Height          =   1185
         Left            =   0
         TabIndex        =   4
         Top             =   855
         Width           =   5550
         _cx             =   9790
         _cy             =   2090
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
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillIn.frx":1192
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
         Begin VB.PictureBox picImgIn 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgColIn 
               Height          =   195
               Left            =   0
               Picture         =   "frmBillIn.frx":1335
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   135
      ScaleHeight     =   1665
      ScaleWidth      =   2340
      TabIndex        =   1
      Top             =   3630
      Width           =   2340
      Begin MSComctlLib.ListView lvwMain 
         Height          =   3345
         Left            =   -15
         TabIndex        =   2
         Top             =   180
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   5900
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ʊ������"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   635
      SimpleText      =   $"frmBillIn.frx":1883
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBillIn.frx":18CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":215E
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":25B0
            Key             =   "C2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":28CA
            Key             =   "C3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":2BE4
            Key             =   "C5"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":2EFE
            Key             =   "C1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":3218
            Key             =   "C4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":3532
            Key             =   "C7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillIn.frx":384C
            Key             =   "C6"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   390
      Top             =   1605
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmBillIn.frx":3B66
      Left            =   930
      Top             =   1815
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBillIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************************************************************
'����:Ʊ��������
'����:���˺�
'����:2010-11-15 11:38:31
'˵��:
'     33725
'********************************************************************************************************************************************
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrCmb As CommandBarComboBox

Private WithEvents mfrmFilter As frmBillInFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mlngModul As Long, mstrPrivs As String
Private mblnFirst As Boolean  '��һ�μ��ش���
Private mstrKey As String    '��һ�εļ�¼
Private mArrFilter As Variant   '����
Private mobjPaneList As Pane
Private mPanSearch As Pane
Private mstrƱ�� As String   '��һ�ε�Ʊ������
Private mstrƱ�ݳ���  As String  'Ʊ�ݳ���

Private Enum mPgIndex
    Pg_������� = 250101
    Pg_���ü�¼ = 250102
End Enum
Private Enum mPaneID
    Pane_List = 1    '��������
    Pane_Search = 2    '����б�
    Pane_BillIn = 3  'ҳ��
    Pane_BillDetails = 4    '��ϸ�б�
End Enum
Private mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Private mintColumn As Integer '
Private mblnDateMoved As Boolean '��ǰʱ�䷶Χ�Ƿ���ת��֮ǰ
 
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitFace = False Then Unload Me: Exit Sub
    Call LoadInDataToRpt
End Sub
Private Sub Form_Load()
    mblnFirst = True: mstrPrivs = gstrPrivs: mlngModul = glngModul
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars '��ʼ�˵���������
    Call InitPanel: Call InitPage: Call InitVsGrid
     Set mArrFilter = mfrmFilter.GetFilterCon
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.���, UserInfo.����)
    End If
    Err.Clear: On Error GoTo 0
End Sub

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_NewItem     '�������
        Call BillInAdd
    Case conMenu_Edit_Modify     '�޸����
        Call BillInModify
    Case conMenu_Edit_Delete      'ɾ�����
        Call DeleteBillIn
    Case conMenu_Edit_DamnifyAdd      '���ӱ���
        Call ExcuteDamnifyFun(1)
    Case conMenu_Edit_DamnifyDelete      'ɾ������
        Call ExcuteDamnifyFun(2)
    Case conMenu_View_Refresh   'ˢ��
        '����ˢ������
        Call LoadInDataToRpt
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
        
    Case conMenu_Edit_NewItem     '�������
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify     '�޸����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ��")
        With vsBillIn
            If .Row > 0 Then
                Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
            Else
                Control.Enabled = Control.Visible And False
            End If
        End With
    Case conMenu_Edit_Delete      'ɾ�����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ɾ��Ʊ��")
        With vsBillIn
            If .Row > 0 Then
                Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
            Else
                Control.Enabled = Control.Visible And False
            End If
        End With
    Case conMenu_Edit_DamnifyAdd      '���ӱ���
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "���ӱ���")
        With vsBillIn
            If .Row > 0 Then
                Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
            Else
                Control.Enabled = Control.Visible And False
            End If
        End With
    Case conMenu_Edit_DamnifyDelete      'ɾ������
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ɾ������")
        If tbPage.Selected.Tag = mPgIndex.Pg_������� Then
            With vsDamnifyBill
                If .Row > 0 Then
                    Control.Enabled = Control.Visible And Val(.RowData(.Row)) > 0
                Else
                    Control.Enabled = Control.Visible And False
                End If
            End With
        Else
                Control.Enabled = Control.Visible And False
        End If
    Case conMenu_View_Refresh   'ˢ��
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
           ' Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1504" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.ID
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Parameter     '��������
        Case conMenu_Edit_UserType  'ʹ�����ѡ��
            Call SetDefaultUserType: Call LoadInDataToRpt
        Case Else   '�����������ܵ���
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, intƱ�� As gBillType
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case conMenu_Edit_UserType
        If lvwMain.SelectedItem Is Nothing Then
             intƱ�� = 0
         Else
             intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨,6-���ѿ�
         End If
        Select Case intƱ��
        Case gBillType.�շ��վ�, gBillType.�����վ�
            Control.Visible = True
        Case gBillType.Ԥ���վ�
            Control.Visible = True
        Case gBillType.���￨, gBillType.���ѿ�
            Control.Visible = True
        Case Else
            Control.Visible = False
        End Select
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub
Private Sub Form_Initialize()
  Call InitCommonControls
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTemp As String
    Err = 0: On Error Resume Next
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter
    Set mfrmFilter = Nothing
    
   SaveWinState Me, App.ProductName
   zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Name, "����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
   zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Name, "�����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
   zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Name, "�����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
   zlSaveDockPanceToReg Me, dkpMan, "����"
   Set mcbrCmb = Nothing
   mstrƱ�� = ""
   
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = False
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Sub

Private Sub imgColDamnifyBill_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picColDamnifyBill.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + imgColDamnifyBill.Height
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsDamnifyBill, lngLeft, lngTop, imgColDamnifyBill.Height)
    zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Caption, "�����б�", True
End Sub

Private Sub mfrmFilter_WindowsHeight(lngHeght As Long)
    Dim lngHeightY As Long
    lngHeightY = lngHeght / Screen.TwipsPerPixelY
    If dkpMan.FindPane(mPaneID.Pane_Search) Is Nothing Then Exit Sub
    dkpMan.FindPane(mPaneID.Pane_Search).MinTrackSize.Height = lngHeightY
    dkpMan.FindPane(mPaneID.Pane_Search).MaxTrackSize.Height = lngHeightY
    dkpMan.RecalcLayout
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    '���¼�������
    Call LoadInDataToRpt
End Sub

Private Sub picColDamnifyBill_Click()
    Call imgColDamnifyBill_Click
End Sub

 
Private Sub imgColDrawBill_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picColDrawBill.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picColDrawBill.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsDrawBill, lngLeft, lngTop, imgColDrawBill.Height)
    zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Caption, "�����б�", True
End Sub

Private Sub picColDrawBill_Click()
    Call imgColDrawBill_Click
End Sub

Private Sub imgColIn_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgIn.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgIn.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsBillIn, lngLeft, lngTop, imgColIn.Height)
    zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Caption, "����б�", True
End Sub

Private Sub picImgIn_Click()
    Call imgColIn_Click
End Sub

Private Sub lvwMain_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvwMain.Drag 0
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrƱ�� = Item.Key Then Exit Sub
    mstrƱ�� = Item.Key
    
    '�����б�����ʾ����
    Call UpdateGridColumnCaption(CurrentIsBill(Val(Mid(Item.Key, 2))))
    
    Call LoadCombox(mcbrCmb)
    Call LoadInDataToRpt '���¼�������
End Sub

Private Sub UpdateGridColumnCaption(ByVal blnBill As Boolean)
    '������ʾ����
    On Error GoTo ErrHandler
    With vsBillIn
        .TextMatrix(0, .ColIndex("��ʼƱ��")) = IIf(blnBill, "��ʼƱ��", "��ʼ����")
        .TextMatrix(0, .ColIndex("��ֹƱ��")) = IIf(blnBill, "��ֹƱ��", "��ֹ����")
        .TextMatrix(0, .ColIndex("Ʊ������")) = IIf(blnBill, "Ʊ������", "��Ƭ����")
    End With
    With vsDamnifyBill
        .TextMatrix(0, .ColIndex("��ʼƱ��")) = IIf(blnBill, "��ʼƱ��", "��ʼ����")
        .TextMatrix(0, .ColIndex("��ֹƱ��")) = IIf(blnBill, "��ֹƱ��", "��ֹ����")
    End With
    With vsDrawBill
        .TextMatrix(0, .ColIndex("��ʼƱ��")) = IIf(blnBill, "��ʼƱ��", "��ʼ����")
        .TextMatrix(0, .ColIndex("��ֹƱ��")) = IIf(blnBill, "��ֹƱ��", "��ֹ����")
        .TextMatrix(0, .ColIndex("��ǰ����")) = IIf(blnBill, "��ǰ����", "��ǰ����")
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If lvwMain.HitTest(x, y) Is Nothing Then Exit Sub
        lvwMain.Drag 1
    End If
End Sub


Private Sub picBillIn_Resize()
    Err = 0: On Error Resume Next
    With picBillIn
        vsBillIn.Left = .ScaleLeft
        vsBillIn.Top = .ScaleTop
        vsBillIn.Width = .ScaleWidth
        vsBillIn.Height = .ScaleHeight
    End With
End Sub
Private Sub picDamnifyBill_Resize()
    Err = 0: On Error Resume Next
    With picDamnifyBill
        vsDamnifyBill.Left = .ScaleLeft
        vsDamnifyBill.Top = .ScaleTop
        vsDamnifyBill.Width = .ScaleWidth
        vsDamnifyBill.Height = .ScaleHeight
    End With
End Sub
Private Sub picDrawBill_Resize()
    Err = 0: On Error Resume Next
    With picDrawBill
        vsDrawBill.Left = .ScaleLeft
        vsDrawBill.Top = .ScaleTop
        vsDrawBill.Width = .ScaleWidth
        vsDrawBill.Height = .ScaleHeight
    End With
End Sub

 
Private Sub picPage_Resize()
    Err = 0: On Error Resume Next
    With picPage
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picType_Resize()
    Err = 0: On Error Resume Next
    With picType
        lvwMain.Left = .ScaleLeft
        lvwMain.Top = .ScaleTop
        lvwMain.Width = .ScaleWidth
        lvwMain.Height = .ScaleHeight
    End With
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '���:
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-15 11:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�������(&N)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸����(&M)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�����(&D)"):
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyAdd, "���ӱ���(&A)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyDelete, "ɾ������(&R)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("N"), conMenu_Edit_DamnifyAdd
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyAdd, "���ӱ���"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DamnifyDelete, "ɾ������")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
        Set mcbrCmb = .Add(xtpControlComboBox, conMenu_Edit_UserType, "ʹ�����")
        mcbrCmb.Flags = xtpFlagRightAlign
        mcbrCmb.HideFlags = xtpNoHide
        With mcbrCmb
            .Clear
            .Width = (TextWidth("ʹ�����") * 4) / Screen.TwipsPerPixelX
        End With
         mcbrCmb.Style = xtpComboLabel
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_Edit_UserType Then
            mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function LoadCombox(ByVal objCombox As CommandBarComboBox) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Combox����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-27 10:22:29
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intƱ�� As gBillType, strSQL As String, rsTemp As ADODB.Recordset
    Dim str��� As String
    
    On Error GoTo errHandle
    
    If objCombox Is Nothing Then Exit Function
    If lvwMain.SelectedItem Is Nothing Then
        intƱ�� = 0
    Else
        intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨,6-���ѿ�
        str��� = lvwMain.SelectedItem.Tag
    End If
    
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        strSQL = "Select ����,����,����,ȱʡ��־ From Ʊ��ʹ����� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With objCombox
            .Clear
            .AddItem "�������"
            If str��� = "�������" Then .ListIndex = .ListCount
            .AddItem " "
            If str��� = " " Then .ListIndex = .ListCount
            
            .ItemData(.ListCount) = -1
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����)
                .ItemData(.ListCount) = 1
                If Val(NVL(rsTemp!ȱʡ��־)) = 1 And .ListIndex <= 0 Then .ListIndex = .ListCount
                If str��� = NVL(rsTemp!����) Then .ListIndex = .ListCount
                rsTemp.MoveNext
            Loop
            If .ListIndex <= 0 Then .ListIndex = 1
        End With
        mcbrCmb.Caption = "ʹ�����"
   Case gBillType.Ԥ���վ�
        With objCombox
            .Clear
            If InStr(1, mstrPrivs, ";Ԥ������Ʊ��;") > 0 _
                And InStr(1, mstrPrivs, ";Ԥ��סԺƱ��;") > 0 Then
                .AddItem "����Ԥ��"
                .ItemData(.ListCount) = 0
                If Val(str���) = 0 Then .ListIndex = 0
            End If
            If InStr(1, mstrPrivs, ";Ԥ������Ʊ��;") > 0 Then
                .AddItem "����Ԥ��": .ItemData(.ListCount) = 2
                If Val(str���) = 2 Then .ListIndex = .ListCount
            End If
            If InStr(1, mstrPrivs, ";Ԥ��סԺƱ��;") > 0 Then
                .AddItem "סԺԤ��": .ItemData(.ListCount) = 3
                If Val(str���) = 3 Then .ListIndex = .ListCount
            End If
            '58071
            If InStr(1, mstrPrivs, ";Ԥ������Ʊ��;") > 0 _
                And InStr(1, mstrPrivs, ";Ԥ��סԺƱ��;") > 0 Then
                .AddItem " ": .ItemData(.ListCount) = 1
                If Val(str���) = 1 Then .ListIndex = .ListCount
            End If
            If .ListIndex <= 0 And .ListCount > 0 Then .ListIndex = 1
        End With
        mcbrCmb.Caption = "ʹ�����"
    Case gBillType.���￨
        strSQL = "Select ID,����,����,ȱʡ��־ From ҽ�ƿ���� where nvl(�Ƿ�����,0) >=1 Order by ���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With objCombox
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
                .ItemData(.ListCount) = Val(NVL(rsTemp!ID))
                If Val(NVL(rsTemp!ȱʡ��־)) = 1 And .ListIndex <= 0 Then .ListIndex = .ListCount
                If Val(str���) = Val(NVL(rsTemp!ID)) Then .ListIndex = .ListCount
                rsTemp.MoveNext
            Loop
            If .ListIndex <= 0 And .ListCount > 0 Then .ListIndex = .ListCount
        End With
        mcbrCmb.Caption = "�����"
    Case gBillType.���ѿ�
        strSQL = "Select ���,���� From ���ѿ����Ŀ¼ where nvl(����,0) >=1 Order by ��� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With objCombox
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!����)
                .ItemData(.ListCount) = Val(NVL(rsTemp!���))
                If Val(str���) = Val(NVL(rsTemp!���)) Then .ListIndex = .ListCount
                rsTemp.MoveNext
            Loop
            If .ListIndex <= 0 And .ListCount > 0 Then .ListIndex = 1
        End With
        mcbrCmb.Caption = "�����"
    Case Else
    End Select
    LoadCombox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetSearchWindowsHeight()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��Ҵ���ĸ߶�
    '����:���˺�
    '����:2010-11-15 14:50:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngSearchHeight As Long
    lngSearchHeight = mfrmFilter.Height / Screen.TwipsPerPixelY
    mPanSearch.MinTrackSize.Height = lngSearchHeight: mPanSearch.MaxTrackSize.Height = lngSearchHeight
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2010-11-15 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, lngWidth As Long
    Dim lngHeight As Long
    
    If mfrmFilter Is Nothing Then
        Set mfrmFilter = New frmBillInFilter
        Load mfrmFilter
    End If
    lngWidth = lvwMain.Width / Screen.TwipsPerPixelX
    With dkpMan
        Set mobjPaneList = .CreatePane(mPaneID.Pane_List, lngWidth, 400, DockLeftOf, Nothing)
        mobjPaneList.Title = "Ʊ����Ϣ": mobjPaneList.Options = PaneNoCloseable
        mobjPaneList.MinTrackSize.Width = lngWidth: mobjPaneList.MaxTrackSize.Width = lngWidth
        mobjPaneList.Handle = picType.hWnd
        mobjPaneList.Tag = mPaneID.Pane_List
        
        Set mPanSearch = .CreatePane(mPaneID.Pane_Search, 400, 400, DockRightOf, mobjPaneList)
        mPanSearch.Title = "��������"
        mPanSearch.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        mPanSearch.Handle = mfrmFilter.hWnd
        mPanSearch.Tag = mPaneID.Pane_Search
        
        Set objPane = .CreatePane(mPaneID.Pane_BillIn, 400, 400, DockBottomOf, mPanSearch)
        objPane.Title = "Ʊ�������Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBillIn.hWnd
        objPane.Tag = mPaneID.Pane_BillIn
        Set objPane = .CreatePane(mPaneID.Pane_BillDetails, 400, 400, DockBottomOf, objPane)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picPage.hWnd
        objPane.Tag = mPaneID.Pane_BillDetails
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
         
    End With
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Function
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPaneID.Pane_Search    '������������
        Item.Handle = mfrmFilter.hWnd
    Case mPaneID.Pane_BillIn    '�����Ϣ
        Item.Handle = picBillIn.hWnd
    Case mPaneID.Pane_BillDetails  '��ϸ�б�
        Item.Handle = picPage.hWnd
    Case mPaneID.Pane_List
        Item.Handle = picType.hWnd
    End Select
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2010-11-15 13:56:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�������, "�������", picDamnifyBill.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_�������
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_���ü�¼, "������Ϣ", picDrawBill.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_���ü�¼
     With tbPage
        tbPage.Item(i).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2010-11-15 13:58:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBillIn
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("�������")) = "1|0"
        .ColData(.ColIndex("��ʼƱ��")) = "1|1"
        .ColData(.ColIndex("��ֹƱ��")) = "1|0"
    End With
    With vsDamnifyBill
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("��������")) = "1|0"
        .ColData(.ColIndex("��ʼƱ��")) = "1|1"
        .ColData(.ColIndex("��ֹƱ��")) = "1|0"
    End With
    With vsDrawBill
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("��ʼƱ��")) = "1|1"
        .ColData(.ColIndex("��ֹƱ��")) = "1|0"
    End With
End Sub
Private Function LoadInDataToRpt() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ݸ�����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-15 14:54:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, intƱ�� As gBillType, lngPreID As Long
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim str��� As String
    Err = 0: On Error GoTo ErrHand:
    
    If lvwMain.SelectedItem Is Nothing Then
        intƱ�� = 0
    Else
        intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨,6-���ѿ�
    End If
    strWhere = ""
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        str��� = mcbrCmb.Text
        If str��� = " " Then
              strWhere = strWhere & " And nvl(A.ʹ�����,'LXH')=[6]"
              str��� = "LXH"
        ElseIf str��� <> "�������" Then
              strWhere = strWhere & " And nvl(A.ʹ�����,'LXH')=[6]"
        End If
    Case gBillType.Ԥ���վ�
        If mcbrCmb.ListIndex <= 0 Then Exit Function
        str��� = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        If Val(str���) <> 0 Then
            str��� = Val(str���) - 1
            strWhere = strWhere & " And nvl(A.ʹ�����,'0')=[6]"
        End If
    Case gBillType.���￨
        If mcbrCmb.ListIndex <= 0 Then Exit Function
        str��� = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        If Val(str���) <> 0 Then
            str��� = Val(str���)
            strWhere = strWhere & " And nvl(A.ʹ�����,'0')=[6]"
        End If
    Case gBillType.���ѿ�
        If mcbrCmb.ListIndex <= 0 Then Exit Function
        str��� = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        If Val(str���) <> 0 Then
            str��� = Val(str���)
            strWhere = strWhere & " And nvl(A.�ӿڱ��,0)=[6]"
        End If
    Case Else
    End Select
    
    
    If mArrFilter("���ʱ��")(0) <> "1901-01-01" And mArrFilter("���ʱ��")(0) <> "1901-01-01" Then
        strWhere = strWhere & " And (A.�Ǽ�ʱ�� Between [1] And [2] )"
    End If
    
    If mArrFilter("�Ǽ���") <> "" Then
        strWhere = strWhere & " And A.�Ǽ���=[3]"
    End If
    If Val(mArrFilter("����ʾ�п�淢Ʊ")) = 1 Then
        If intƱ�� = gBillType.���ѿ� Then
            strWhere = strWhere & " And nvl(A.ʣ������,0)>0 and A.�Ƿ���ڿ� =1"
        Else
            strWhere = strWhere & " And nvl(A.ʣ������,0)>0 and A.����Ʊ�� =1"
        End If
    End If
    If intƱ�� <> gBillType.���ѿ� Then
        strWhere = strWhere & " And nvl(A.Ʊ��,0)=[4]"
    End If
    If InStr(1, mstrPrivs, ";����������˵Ǽ�Ʊ��;") = 0 Then
        strWhere = strWhere & " And A.�Ǽ��� =[5]"
    End If
    If intƱ�� = gBillType.���ѿ� Then
        gstrSQL = "" & _
        "  Select A.Id,nvl(B.����,'���￨') as ʹ�����, A.ǰ׺�ı�, A.��ʼ���� As ��ʼ����,A.��ֹ���� As ��ֹ����, A.�������," & _
        "       A.ʣ������, A.��ע, A.�Ǽ���, A.�Ǽ�ʱ��, A.����  " & vbCrLf & _
        "  From ���ѿ�����¼ A,���ѿ����Ŀ¼ B" & vbCrLf & _
        "  Where nvl(A.�ӿڱ��,0)=B.���(+) And " & Mid(strWhere, 5)
    ElseIf intƱ�� = gBillType.���￨ Then
        gstrSQL = "" & _
        "  Select A.Id,nvl(B.����,'���￨') as ʹ�����, A.ǰ׺�ı�, A.��ʼ����,A.��ֹ����, A.�������," & _
        "       A.ʣ������, A.��ע, A.�Ǽ���, A.�Ǽ�ʱ��, A.����  " & vbCrLf & _
        "  From Ʊ������¼ A,ҽ�ƿ���� B" & vbCrLf & _
        "  where   to_number(nvl(A.ʹ�����,'0'))=B.ID(+) And " & Mid(strWhere, 5)
    ElseIf intƱ�� = gBillType.Ԥ���վ� Then
        gstrSQL = "" & _
        "  Select A.Id,decode(nvl(A.ʹ�����,0),0,'','1','����','סԺ') as ʹ�����, A.Ʊ��, A.ǰ׺�ı�, A.��ʼ����,A.��ֹ����, A.�������," & _
        "       A.ʣ������, A.��ע, A.�Ǽ���, A.�Ǽ�ʱ��, A.����  " & vbCrLf & _
        "  From Ʊ������¼ A" & vbCrLf & _
        "  where   " & Mid(strWhere, 5)
    Else
        gstrSQL = "" & _
        "  Select A.Id,A.ʹ�����, A.ǰ׺�ı�, A.��ʼ����,A.��ֹ����, A.�������," & _
        "       A.ʣ������, A.��ע, A.�Ǽ���, A.�Ǽ�ʱ��, A.����  " & vbCrLf & _
        "  From Ʊ������¼ A" & vbCrLf & _
        "  where   " & Mid(strWhere, 5)
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("���ʱ��")(0)), CDate(mArrFilter("���ʱ��")(1)), _
        CStr(mArrFilter("�Ǽ���")), intƱ��, UserInfo.����, str���)
    
    With vsBillIn
        .Clear 1
        If .Row > 0 Then
            lngPreID = Val(.RowData(.Row))
        End If
        .Redraw = flexRDNone: .Clear 1: .Rows = 2
        .RowData(1) = 0
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!ʹ�����)
            .TextMatrix(lngRow, .ColIndex("�������")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("ǰ׺�ı�")) = NVL(rsTemp!ǰ׺�ı�)
            .TextMatrix(lngRow, .ColIndex("��ʼƱ��")) = NVL(rsTemp!��ʼ����)
            .TextMatrix(lngRow, .ColIndex("��ֹƱ��")) = NVL(rsTemp!��ֹ����)
            .TextMatrix(lngRow, .ColIndex("Ʊ������")) = NVL(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("ʣ������")) = NVL(rsTemp!ʣ������)
            .TextMatrix(lngRow, .ColIndex("�Ǽ���")) = NVL(rsTemp!�Ǽ���)
            .TextMatrix(lngRow, .ColIndex("�Ǽ�ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("��ע")) = NVL(rsTemp!��ע)
            If lngPreID = Val(NVL(rsTemp!ID)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .FixedCols = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBillIn, Me.Name, "����б�", True, True
    Call vsBillIn_AfterRowColChange(-1, 0, vsBillIn.Row, 0)
    vsBillIn.ColHidden(vsBillIn.ColIndex("��־")) = False
    vsBillIn.ColWidth(vsBillIn.ColIndex("��־")) = 300
    LoadInDataToRpt = True
    Exit Function
ErrHand:
    vsBillIn.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub LoadDamnifyData(ByVal lng������� As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ر�������
    '����:���˺�
    '����:2010-11-15 17:42:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngPreID As Long, lngRow As Long
    Dim intƱ�� As gBillType
    
    On Error GoTo ErrHand:
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If intƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select ID, ���id, ��ʼ���� As ��ʼ����, ��ֹ���� As ��ֹ����, ����, ����ԭ��, ������, ����ʱ��" & vbNewLine & _
            "From ���ѿ������¼" & vbNewLine & _
            "Where ���id = [1]"
    Else
        gstrSQL = _
            "Select ID, ���id, ��ʼ����, ��ֹ����, ����, ����ԭ��, ������, ����ʱ��" & vbNewLine & _
            "From Ʊ�ݱ����¼" & vbNewLine & _
            "Where ���id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�������)
    
    With vsDamnifyBill
        If .Row > 0 Then
            lngPreID = Val(.RowData(.Row))
        End If
        .Redraw = flexRDNone: .Clear 1: .Rows = 2
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .RowData(1) = ""
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("��ʼƱ��")) = NVL(rsTemp!��ʼ����)
            .TextMatrix(lngRow, .ColIndex("��ֹƱ��")) = NVL(rsTemp!��ֹ����)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ԭ��")) = NVL(rsTemp!����ԭ��)
            If lngPreID = Val(NVL(rsTemp!ID)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDamnifyBill, Me.Name, "�����б�", True, True
    Exit Sub
ErrHand:
    vsDamnifyBill.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub LoadDrawData(ByVal lng������� As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:lng�������-���ID
    '����:���˺�
    '����:2010-11-15 17:31:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngPreID As Long, lngRow As Long
    Dim intƱ�� As gBillType
    
    On Error GoTo ErrHand:
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If intƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select ID,������,ǰ׺�ı�,��ʼ���� As ��ʼ����,��ֹ���� As ��ֹ����," & _
            "       decode(ʹ�÷�ʽ,1,'����','����') as ʹ�÷�ʽ,to_char(�Ǽ�ʱ��,'yyyy-MM-dd') as �Ǽ�ʱ��," & _
            "       �Ǽ���,��ǰ���� As ��ǰ����,ʣ������,����,�˶��� " & _
            "From ���ѿ����ü�¼ " & _
            "Where ����=[1]"
    Else
        gstrSQL = "" & _
        "   Select ID,������,ǰ׺�ı�,��ʼ����,��ֹ����," & _
        "           decode(ʹ�÷�ʽ,1,'����','����') as ʹ�÷�ʽ,to_char(�Ǽ�ʱ��,'yyyy-MM-dd') as �Ǽ�ʱ��," & _
        "           �Ǽ���,��ǰ����,ʣ������,����,�˶��� " & _
        "   From Ʊ�����ü�¼ " & _
        "   Where ����=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�������)
    
   With vsDrawBill
        If .Row > 0 Then
            lngPreID = Val(.RowData(.Row))
        End If
        .Redraw = flexRDNone: .Clear 1: .Rows = 2
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            
            .TextMatrix(lngRow, .ColIndex("��ʼƱ��")) = NVL(rsTemp!��ʼ����)
            .TextMatrix(lngRow, .ColIndex("��ֹƱ��")) = NVL(rsTemp!��ֹ����)
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��ǰ����")) = NVL(rsTemp!��ǰ����)
            .TextMatrix(lngRow, .ColIndex("ʣ������")) = NVL(rsTemp!ʣ������)
            .TextMatrix(lngRow, .ColIndex("�������")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("ʹ�÷�ʽ")) = NVL(rsTemp!ʹ�÷�ʽ)
            .TextMatrix(lngRow, .ColIndex("�Ǽ���")) = NVL(rsTemp!�Ǽ���)
            .TextMatrix(lngRow, .ColIndex("�Ǽ�ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("�˶���")) = NVL(rsTemp!�˶���)
            .TextMatrix(lngRow, .ColIndex("ǰ׺�ı�")) = NVL(rsTemp!ǰ׺�ı�)
            If lngPreID = Val(NVL(rsTemp!ID)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDrawBill, Me.Name, "�����б�", True, True
    Exit Sub
ErrHand:
    vsDrawBill.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function InitFace() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2010-11-15 15:13:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrTemp1 As Variant, arrTemp2 As Variant, arrTemp3 As Variant
    Dim i As Integer, strTmp As String
    
    mstrƱ�ݳ��� = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    If zlStr.IsHavePrivs(mstrPrivs, "�շ��վ�") Then
        strTmp = strTmp & "|" & "�շ��վ�,1"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "Ԥ���վ�") _
        And (zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") _
            Or zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��")) Then
        strTmp = strTmp & "|" & "Ԥ���վ�,2"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "�����վ�") Then
        strTmp = strTmp & "|" & "�����վ�,3"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "�Һ��վ�") Then
        strTmp = strTmp & "|" & "�Һ��վ�,4"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "ҽ�ƿ�") Then
        strTmp = strTmp & "|" & "ҽ�ƿ�,5"
    End If
    If zlStr.IsHavePrivs(mstrPrivs, "���ѿ�") Then
        strTmp = strTmp & "|" & "���ѿ�,6"
    End If
    If strTmp = "" Then
        MsgBox "��û�в����κ�Ʊ�ݵ�Ȩ��!", vbInformation, App.ProductName
        Exit Function
    Else
        strTmp = Mid(strTmp, 2)
    End If
    arrTemp1 = Split(strTmp, "|")
    For i = 0 To UBound(arrTemp1)
        arrTemp2 = Split(arrTemp1(i), ",")
        lvwMain.ListItems.Add , "C" & arrTemp2(1), arrTemp2(0), "C" & arrTemp2(1)
    Next
    lvwMain.ListItems(1).Selected = True
    Call LoadCombox(mcbrCmb)
    Call SetDefaultUserType
    InitFace = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetDefaultUserType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ��ʹ�����
    '����:���˺�
    '����:2011-04-27 14:23:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intƱ�� As gBillType
    
    If mcbrCmb Is Nothing Then Exit Sub
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))  '1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨,6-���ѿ�
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        lvwMain.SelectedItem.Tag = mcbrCmb.Text
    Case gBillType.Ԥ���վ�, gBillType.���￨, gBillType.���ѿ�
        If mcbrCmb.ListIndex <= 0 Then
            lvwMain.SelectedItem.Tag = ""
        Else
            lvwMain.SelectedItem.Tag = mcbrCmb.ItemData(mcbrCmb.ListIndex)
        End If
    Case Else
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:
    '����:
    '����:���˺�
    '����:2010-11-15 15:34:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String, bln��ϸ As Boolean
    bln��ϸ = False
    If Me.ActiveControl Is vsDrawBill Then
        Set vsBill = vsDrawBill: strTittle = GetUnitName & lvwMain.SelectedItem.Text & "������ϸ": bln��ϸ = True
    ElseIf Me.ActiveControl Is vsDamnifyBill Then
        Set vsBill = vsDamnifyBill: strTittle = GetUnitName & lvwMain.SelectedItem.Text & "������ϸ": bln��ϸ = True
    Else
        Set vsBill = vsBillIn: strTittle = GetUnitName & lvwMain.SelectedItem.Text & "������"
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    If bln��ϸ Then
        With vsBillIn
            objRow.Add "�������:" & Val(.TextMatrix(.Row, .ColIndex("�������")))
            objRow.Add "Ʊ�ݷ�Χ:" & Trim(.TextMatrix(.Row, .ColIndex("��ʼƱ��"))) & "��" & Trim(.TextMatrix(.Row, .ColIndex("��ֹƱ��")))
            objRow.Add "�Ǽ���:" & Trim(.TextMatrix(.Row, .ColIndex("�Ǽ���")))
        End With
    Else
        If CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
            objRow.Add "���ʱ�䣺" & CStr(mArrFilter("���ʱ��")(0)) & "��" & CStr(mArrFilter("���ʱ��")(1))
        End If
        If Val(mArrFilter("����ʾ�п�淢Ʊ")) = 1 Then
            objRow.Add "����ʾ�п�淢Ʊ"
        End If
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
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
    '�ָ�
    With vsBill
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub ExcuteDamnifyFun(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    ' ���:bytFun-1-����;2-ɾ��
    '����:���˺�
    '����:2010-11-15 15:36:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTemp As String, lngRow As Long
    Dim rsTemp As ADODB.Recordset
    Dim lng���ID As Long, intƱ�� As gBillType
    
    On Error GoTo errHandle
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    With vsBillIn
        lngID = Val(.RowData(.Row))
        lng���ID = lngID
        If lngID < 0 Then Exit Sub
    End With
    If bytFun = 1 Then '����
        If frmBillInDamnifyEdit.zlBillEdit(Me, EdS_����, mstrPrivs, mlngModul, intƱ��, lngID) = False Then Exit Sub
        gstrSQL = "Select ʣ������ From Ʊ������¼ where ID=[1]"
        If intƱ�� = gBillType.���ѿ� Then
            gstrSQL = "Select ʣ������ From ���ѿ�����¼ where ID=[1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng���ID)
        If Not rsTemp.EOF Then
            vsBillIn.TextMatrix(vsBillIn.Row, vsBillIn.ColIndex("ʣ������")) = NVL(rsTemp!ʣ������)
        End If
        Call LoadDamnifyData(lngID)
        Exit Sub
    End If
    
    With vsDamnifyBill
        lngID = Val(.RowData(.Row))
        If lngID = 0 Then Exit Sub
        If Not (.TextMatrix(.Row, .ColIndex("��ֹƱ��")) = "" Or .TextMatrix(.Row, .ColIndex("��ʼƱ��")) = .TextMatrix(.Row, .ColIndex("��ֹƱ��"))) Then
            strTemp = "-" & .TextMatrix(.Row, .ColIndex("��ֹƱ��"))
        End If
        If MsgBox("��ȷ��Ҫɾ�������Ϊ��" & .TextMatrix(.Row, .ColIndex("��ʼƱ��")) & strTemp & "����" & _
             "�����¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End With
    Err = 0: On Error GoTo errHandle:
    
    Me.MousePointer = 11
    If intƱ�� = gBillType.���ѿ� Then
        'Zl_���ѿ������¼_Delete(Id_In In ���ѿ������¼.ID%Type) Is
        gstrSQL = "Zl_���ѿ������¼_Delete(" & lngID & ")"
    Else
        'Zl_Ʊ�ݱ����¼_Delete(Id_In In Ʊ�ݱ����¼.ID%Type) Is
        gstrSQL = "Zl_Ʊ�ݱ����¼_Delete(" & lngID & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Me.MousePointer = 0
    With vsDamnifyBill
        lngRow = .Row
        If .Rows - 1 > .Row Then
            .Row = .Row + 1
            .RemoveItem lngRow
        ElseIf .Row = 1 Then
            .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            .RowData(.Row) = ""
        Else
              .Row = .Row - 1
              .RemoveItem lngRow
        End If
    End With
    gstrSQL = "Select ʣ������ From Ʊ������¼ where ID=[1]"
    If intƱ�� = gBillType.���ѿ� Then
        gstrSQL = "Select ʣ������ From ���ѿ�����¼ where ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng���ID)
    If Not rsTemp.EOF Then
        vsBillIn.TextMatrix(vsBillIn.Row, vsBillIn.ColIndex("ʣ������")) = NVL(rsTemp!ʣ������)
    End If
    
    'If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Private Sub DeleteBillIn()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ�����Ʊ��
    '����:���˺�
    '����:2010-11-15 16:14:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngID As Long
    Dim strExpended As String, intƱ�� As gBillType, blnTrans As Boolean
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    With vsBillIn
        If .Row <= 0 Then Exit Sub
        lngID = Val(.RowData(.Row))
        If lngID < 0 Then Exit Sub
        If MsgBox("��ȷ��Ҫɾ��" & lvwMain.SelectedItem.Text & "������Ϊ��" & .TextMatrix(.Row, .ColIndex("�������")) & "��;��ʼ" & _
            IIf(CurrentIsBill(intƱ��), "����", "����") & "Ϊ��" & .TextMatrix(.Row, .ColIndex("��ʼƱ��")) & "����" & _
            "����¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End With
    '102996:���ϴ�,2016/11/23,ҽ�Ʒ�Ʊ���ӻ�����
    If gblnBillPrint Then
        On Error Resume Next
        If gobjBillPrint.zlBillInCheckValied(3, intƱ��, "", lngID, "", "", strExpended) = False Then Exit Sub
    End If
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    If intƱ�� = gBillType.���ѿ� Then
        'Zl_���ѿ�����¼_Delete(Id_In In ���ѿ�����¼.ID%Type) Is
        gstrSQL = "Zl_���ѿ�����¼_Delete(" & lngID & ")"
    Else
        'Zl_Ʊ������¼_Delete(Id_In In Ʊ������¼.ID%Type) Is
        gstrSQL = "Zl_Ʊ������¼_Delete(" & lngID & ")"
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    If gblnBillPrint Then
        On Error Resume Next
        strExpended = ""
        If gobjBillPrint.zlBillIn(3, intƱ��, "", lngID, strExpended) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            Me.MousePointer = 0
            Exit Sub
        End If
        Err = 0: On Error GoTo errHandle
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    
    Me.MousePointer = 0
    With vsBillIn
        lngRow = .Row
        If .Rows - 1 > .Row Then
            .Row = .Row + 1
            .RemoveItem lngRow
        ElseIf .Row = 1 Then
            .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            .RowData(.Row) = ""
        Else
              .Row = .Row - 1
              .RemoveItem lngRow
        End If
    End With
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
    Me.MousePointer = 0
End Sub
Private Sub BillInModify()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸����Ʊ��
    '����:���˺�
    '����:2010-11-15 17:01:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID  As Long, strNO As String, rsTemp As ADODB.Recordset, lngLen As Long
    Dim intƱ�� As gBillType
    
    With vsBillIn
        If .Row <= 0 Then Exit Sub
        lngID = Val(.RowData(.Row))
        If lngID < 0 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("��ʼƱ��")))
    End With
    '102181:���ϴ�,2016/11/10,ҽ�ƿ�Ʊ�ݳ���
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If Not (intƱ�� = gBillType.���￨ Or intƱ�� = gBillType.���ѿ�) Then
        lngLen = Val(Split(mstrƱ�ݳ���, "|")(Mid(lvwMain.SelectedItem.Key, 2) - 1))
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ��") = False Then
        If frmBillInEdit.zlBillEdit(Me, intƱ��, Ed_�鿴, mstrPrivs, mlngModul, lngID) = False Then Exit Sub
        Exit Sub
    End If
 
    If frmBillInEdit.zlBillEdit(Me, intƱ��, Ed_�޸�, mstrPrivs, mlngModul, lngID) = False Then Exit Sub
    '    If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call LoadInDataToRpt
End Sub
Private Sub BillInAdd()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2010-11-15 17:01:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intƱ�� As gBillType
    Dim str��� As String
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    If zlStr.IsHavePrivs(mstrPrivs, "����Ʊ��") = False Then Exit Sub
    
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If intƱ�� = gBillType.�շ��վ� Or intƱ�� = gBillType.�����վ� Then
        If Not mcbrCmb Is Nothing Then str��� = Trim(mcbrCmb.Text)
    ElseIf intƱ�� = gBillType.Ԥ���վ� Or intƱ�� = gBillType.���￨ Or intƱ�� = gBillType.���ѿ� Then
        If Not mcbrCmb Is Nothing Then
            If mcbrCmb.ListIndex > 0 Then
                str��� = Trim(mcbrCmb.ItemData(mcbrCmb.ListIndex))
            End If
        End If
    End If
    
    If frmBillInEdit.zlBillEdit(Me, intƱ��, Ed_����, mstrPrivs, mlngModul, 0, str���) = False Then Exit Sub
    '    If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call LoadInDataToRpt
End Sub

Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode������
    '����:���˺�
    '����:2010-11-15 17:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�Ǽ��� As String, str���ID As String, intƱ�� As Integer
    
    With vsBillIn
        If .Row > 0 Then
            str���ID = Val(.RowData(.Row))
            str�Ǽ��� = Trim(.TextMatrix(.Row, .ColIndex("�Ǽ���")))
        End If
    End With
    intƱ�� = Val(Mid(lvwMain.SelectedItem.Key, 2))
 
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, _
        "Ʊ��=" & intƱ��, "�Ǽ���=" & str�Ǽ���, "���ID=" & str���ID)
End Sub

Private Sub vsBillIn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngID As Long
    If OldRow = NewRow Then Exit Sub
    With vsBillIn
        lngID = Val(.RowData(NewRow))
        If lngID = 0 Then lngID = -1 '������ʾ��������̵����ü�¼
        Call LoadDrawData(lngID)
        Call LoadDamnifyData(lngID)
    End With
End Sub

Private Sub vsBillIn_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillIn
        If .ColIndex("��־") = Col Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsBillIn_DblClick()
    Call BillInModify
End Sub
Private Sub vsBillIn_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Call BillInModify
End Sub
Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ��������
    '����:���˺�
    '����:2010-11-15 17:54:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.ActiveControl Is vsDrawBill Then
        With vsDrawBill
            If .Rows < 1 Then Exit Function
            zlIsHaveData = Val(.RowData(1)) > 0: Exit Function
        End With
    ElseIf Me.ActiveControl Is vsDamnifyBill Then
        With vsDrawBill
            If .Rows < 1 Then Exit Function
            zlIsHaveData = Val(.RowData(1)) > 0: Exit Function
        End With
    Else
        With vsBillIn
            If .Rows < 0 Then Exit Function
            zlIsHaveData = Val(.RowData(1)) > 0: Exit Function
        End With
    End If
End Function

Private Sub vsBillIn_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Name, "����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBillIn_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBillIn, Me.Name, "����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsDamnifyBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDamnifyBill
        If .ColIndex("��־") = Col Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsDrawBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Name, "�����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsDrawBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDrawBill, Me.Name, "�����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsDamnifyBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Name, "�����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsDamnifyBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDamnifyBill, Me.Name, "�����б�", True, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
 
Private Sub vsDrawBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDrawBill
        If .ColIndex("��־") = Col Then Cancel = True: Exit Sub
    End With
End Sub
