VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "CODEJOCK.CALENDAR.V16.3.1.OCX"
Begin VB.Form frmServiceApp 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picReg 
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   2115
      ScaleHeight     =   5280
      ScaleWidth      =   8115
      TabIndex        =   25
      Top             =   525
      Width           =   8115
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   75
         Width           =   1125
      End
      Begin VB.PictureBox picSplit 
         Height          =   50
         Left            =   15
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   4065
         TabIndex        =   54
         Top             =   3105
         Width           =   4065
      End
      Begin VB.CommandButton cmdDirectApp 
         Height          =   315
         Left            =   5295
         Picture         =   "frmServiceApp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   68
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   4380
         TabIndex        =   43
         Top             =   75
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   163971074
         CurrentDate     =   42340
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Left            =   465
         TabIndex        =   41
         ToolTipText     =   "����ͨ���������,ҽ��,����,��Ŀ���Ƽ��������п��ٹ��˲���"
         Top             =   68
         Width           =   1320
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2415
         Left            =   60
         TabIndex        =   46
         Top             =   4555
         Width           =   5925
         _cx             =   10451
         _cy             =   4260
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceApp.frx":058A
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
         Editable        =   2
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
         Height          =   2415
         Left            =   60
         TabIndex        =   45
         Top             =   450
         Width           =   3360
         _cx             =   5927
         _cy             =   4260
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   15658734
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   322
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         ExplorerBar     =   1
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "ʱ���"
         Height          =   180
         Left            =   1875
         TabIndex        =   55
         Top             =   135
         Width           =   540
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼʱ��"
         Height          =   180
         Left            =   3630
         TabIndex        =   47
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   60
         TabIndex        =   40
         Top             =   135
         Width           =   360
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   1065
      ScaleHeight     =   2550
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   405
      Width           =   4845
      Begin VB.TextBox txtMarriage 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2130
         Width           =   1710
      End
      Begin VB.TextBox txtJob 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2130
         Width           =   1710
      End
      Begin VB.TextBox txtNation 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1725
         Width           =   1710
      End
      Begin VB.TextBox txtRace 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1725
         Width           =   1710
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1335
         Width           =   4230
      End
      Begin VB.TextBox txtPhone 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   930
         Width           =   1710
      End
      Begin VB.TextBox txtFeeType 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   930
         Width           =   1710
      End
      Begin VB.TextBox txtBirth 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   525
         Width           =   1710
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   525
         Width           =   690
      End
      Begin VB.TextBox txtGender 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   465
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   1710
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   555
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   1710
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "����״��"
         Height          =   180
         Left            =   2340
         TabIndex        =   23
         Top             =   2205
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   165
         TabIndex        =   21
         Top             =   2205
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2700
         TabIndex        =   17
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "��סַ"
         Height          =   180
         Left            =   -15
         TabIndex        =   15
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��ϵ�绰"
         Height          =   180
         Left            =   2340
         TabIndex        =   13
         Top             =   1005
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   1005
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   2340
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1185
         TabIndex        =   7
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   165
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "���֤��"
         Height          =   180
         Left            =   2340
         TabIndex        =   3
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   360
      End
   End
   Begin VB.PictureBox picApp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   885
      ScaleHeight     =   4245
      ScaleWidth      =   4860
      TabIndex        =   26
      Top             =   2040
      Width           =   4890
      Begin VB.TextBox txtMoney 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox txtRegTime 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   1500
      End
      Begin VB.TextBox txtReger 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   1500
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1650
         Width           =   1500
      End
      Begin VB.TextBox txtDoc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1650
         Width           =   1500
      End
      Begin VB.TextBox txtTimeEnd 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txtTimeBegin 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txtState 
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   90
         Width           =   3915
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "�ҺŽ��"
         Height          =   180
         Left            =   2550
         TabIndex        =   53
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��Ŀ"
         Height          =   180
         Left            =   135
         TabIndex        =   51
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ�ʱ��"
         Height          =   180
         Left            =   2550
         TabIndex        =   39
         Top             =   555
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ���"
         Height          =   180
         Left            =   330
         TabIndex        =   37
         Top             =   555
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����"
         Height          =   180
         Left            =   150
         TabIndex        =   35
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼҽ��"
         Height          =   180
         Left            =   2550
         TabIndex        =   33
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "ʱ�䷶Χ                     ��"
         Height          =   180
         Left            =   150
         TabIndex        =   30
         Top             =   945
         Width           =   2790
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "�Ǽ�ԭ��"
         Height          =   180
         Left            =   150
         TabIndex        =   28
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   600
      ScaleHeight     =   3315
      ScaleWidth      =   4800
      TabIndex        =   48
      Top             =   2985
      Width           =   4800
      Begin XtremeCalendarControl.DatePicker dtpMain 
         Height          =   2895
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   3870
         _Version        =   1048579
         _ExtentX        =   6826
         _ExtentY        =   5106
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowTodayButton =   0   'False
         ShowNoneButton  =   0   'False
         Show3DBorder    =   0
         MaxSelectionCount=   1
         AskDayMetrics   =   -1  'True
      End
   End
   Begin VB.Image imgApp 
      Height          =   240
      Left            =   1860
      Picture         =   "frmServiceApp.frx":0696
      Top             =   645
      Width           =   240
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1230
      Top             =   1515
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnNotClick As Boolean
Private mrsInfo As ADODB.Recordset
Private mlng��ϢID As Long, mfrmMain As Object
Private mstrʱ��s As String, mlngԤԼ��Чʱ��
Private mblnUnload As Boolean, mdatCache As Date
Private mblnFirst As Boolean, mintͬ����Լ�� As Integer
Private mintר�Һ�ԤԼ���� As Integer, mint����ԤԼ������ As Integer
Private mblnInit As Boolean
Private mblnKeyPress As Boolean, mblnAppointmentChange As Boolean
Private mstrPriceGrade As String

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 145, 80, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "���˻�����Ϣ"
    objPane.Handle = picInfo.Hwnd
    objPane.MaxTrackSize.Width = 325
    objPane.MinTrackSize.Width = 325
    objPane.MaxTrackSize.Height = 170
    objPane.MinTrackSize.Height = 170
    
    
    Set objPane = dkpMain.CreatePane(2, 145, 90, DockBottomOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "ԤԼ�Ǽ���Ϣ"
    objPane.Handle = picApp.Hwnd
    objPane.MaxTrackSize.Height = 138
    objPane.MinTrackSize.Height = 138
    
    Set objPane = dkpMain.CreatePane(3, 145, 120, DockBottomOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Handle = picDate.Hwnd
    
    
    Set objPane = dkpMain.CreatePane(4, 320, 400, DockRightOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Handle = picReg.Hwnd
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        .PaintManager.HighlighActiveCaption = False
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function zlGet��ǰ���ڼ�(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���������ڼ�
    '����:���˺�
    '����:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bln��ǰ���� As Boolean, strTemp As String
    If strDate = "" Then
        strSQL = "Select Decode(To_Char(Sysdate,'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��',NULL) as ����  From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "Select Decode(To_Char([1],'D'),'1','��','2','һ','3','��','4','��','5','��','6','��','7','��','') As ���� From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!����)
    zlGet��ǰ���ڼ� = strTemp
End Function

Private Function Check����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As Boolean
'����:�жϲ����Ƿ��ٴε�����ͬ�ٴ����ʵ��ٴ����ҡ��Һ�
'     �����ҹ��ŵ�,��ס��Ժ��,���ﲻ��ȷ��ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select a.�ٴ�����id" & vbNewLine & _
    "       From (Select ִ�в���id �ٴ�����id From ���˹Һż�¼ Where ����id = [1] and ��¼����=1 and ��¼״̬=1 " & vbNewLine & _
    "             Union All" & vbNewLine & _
    "             Select ��Ժ����id �ٴ�����id From ������ҳ Where ����id = [1]) a" & vbNewLine & _
    "       Where Exists (Select 1" & vbNewLine & _
    "                    From �ٴ����� b" & vbNewLine & _
    "                    Where b.����id = a.�ٴ�����id And b.�������� = (Select �������� From �ٴ����� Where ����id = [2] And Rownum=1))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngִ�в���ID)
    Check���� = Not rsTmp.EOF
End Function

Public Function CheckLimit(lng��¼ID As Long) As Boolean
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim rsUsed As ADODB.Recordset, lng������λ�������� As Long
    Dim rsUnit As ADODB.Recordset, lng������λ���� As Long
    
    strSQL = "Select Nvl(��Լ��,�޺���) As ��Լ��,��Լ��,Nvl(�Ƿ��ռ,0) As �Ƿ��ռ,�Ƿ���ſ���,�Ƿ��ʱ�� From �ٴ������¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    strSQL = "Select ���� As ������λ, ���Ʒ�ʽ, ���, ���� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id = [1] And ���� = 1"
    Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    strSQL = "Select Count(1) As ���� From ���˹Һż�¼ Where �����¼id = [1] And ������λ Is Not Null And ��¼״̬=1"
    Set rsUsed = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID)
    If Not rsUsed.EOF Then
        lng������λ�������� = Val(Nvl(rsUsed!����))
    End If
    If Not rsTemp.EOF Then
        If Val(Nvl(rsTemp!�Ƿ���ſ���)) = 1 Then
            If rsUnit.EOF Then
                lng������λ���� = 0
            Else
                If Val(Nvl(rsUnit!���Ʒ�ʽ)) = 2 Then
                    If Val(rsTemp!�Ƿ��ռ) = 0 Then
                        lng������λ���� = 0
                    Else
                        Do While Not rsUnit.EOF
                            lng������λ���� = lng������λ���� + Val(Nvl(rsUnit!����))
                            rsUnit.MoveNext
                        Loop
                    End If
                ElseIf Val(Nvl(rsUnit!���Ʒ�ʽ)) = 1 Then
                    If Val(rsTemp!�Ƿ��ռ) = 0 Then
                        lng������λ���� = 0
                    Else
                        Do While Not rsUnit.EOF
                            lng������λ���� = lng������λ���� + Int(Val(Nvl(rsUnit!����)) * Val(Nvl(rsTemp!��Լ��)) / 100)
                            rsUnit.MoveNext
                        Loop
                    End If
                ElseIf Val(Nvl(rsUnit!���Ʒ�ʽ)) = 3 Then
                    Do While Not rsUnit.EOF
                        lng������λ���� = lng������λ���� + Val(Nvl(rsUnit!����))
                        rsUnit.MoveNext
                    Loop
                End If
            End If
            If Not IsNull(rsTemp!��Լ��) Then
                If Val(Nvl(rsTemp!��Լ��)) + lng������λ���� - lng������λ�������� >= Val(Nvl(rsTemp!��Լ��)) Then
                    MsgBox "��ǰԤԼ���볬������������" & Val(Nvl(rsTemp!��Լ��)) & "(���а���������λ��������" & lng������λ���� & "),���ܼ���ԤԼ!", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If Val(Nvl(rsTemp!�Ƿ��ʱ��)) = 1 Then
                If rsUnit.EOF Then
                    lng������λ���� = 0
                Else
                    If Val(Nvl(rsUnit!���Ʒ�ʽ)) = 3 Then
                        rsUnit.Filter = "���=" & Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col))
                        If rsUnit.EOF Then
                            If Not IsNull(rsTemp!��Լ��) Then
                                If Val(Nvl(rsTemp!��Լ��)) >= Val(Nvl(rsTemp!��Լ��)) Then
                                    MsgBox "��ǰԤԼ���볬������������" & Val(Nvl(rsTemp!��Լ��)) & ",���ܼ���ԤԼ!", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        Else
                            If Not IsNull(rsTemp!��Լ��) Then
                                If Val(Nvl(rsTemp!��Լ��)) >= Val(Nvl(rsTemp!��Լ��)) Then
                                    MsgBox "��ǰԤԼ���볬������������" & Val(Nvl(rsTemp!��Լ��)) & ",���ܼ���ԤԼ!", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If Val(Nvl(rsUnit!���Ʒ�ʽ)) = 2 Then
                            If Val(rsTemp!�Ƿ��ռ) = 0 Then
                                lng������λ���� = 0
                            Else
                                Do While Not rsUnit.EOF
                                    lng������λ���� = lng������λ���� + Val(Nvl(rsUnit!����))
                                    rsUnit.MoveNext
                                Loop
                            End If
                        ElseIf Val(Nvl(rsUnit!���Ʒ�ʽ)) = 1 Then
                            If Val(rsTemp!�Ƿ��ռ) = 0 Then
                                lng������λ���� = 0
                            Else
                                Do While Not rsUnit.EOF
                                    lng������λ���� = lng������λ���� + Int(Val(Nvl(rsUnit!����)) * Val(Nvl(rsTemp!��Լ��)) / 100)
                                    rsUnit.MoveNext
                                Loop
                            End If
                        End If
                        If Not IsNull(rsTemp!��Լ��) Then
                            If Val(Nvl(rsTemp!��Լ��)) + lng������λ���� - lng������λ�������� >= Val(Nvl(rsTemp!��Լ��)) Then
                                MsgBox "��ǰԤԼ���볬������������" & Val(Nvl(rsTemp!��Լ��)) & "(���а���������λ��������" & lng������λ���� & "),���ܼ���ԤԼ!", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Else
                If rsUnit.EOF Then
                    lng������λ���� = 0
                Else
                    If Val(Nvl(rsUnit!���Ʒ�ʽ)) = 2 Then
                        If Val(rsTemp!�Ƿ��ռ) = 0 Then
                            lng������λ���� = 0
                        Else
                            Do While Not rsUnit.EOF
                                lng������λ���� = lng������λ���� + Val(Nvl(rsUnit!����))
                                rsUnit.MoveNext
                            Loop
                        End If
                    ElseIf Val(Nvl(rsUnit!���Ʒ�ʽ)) = 1 Then
                        If Val(rsTemp!�Ƿ��ռ) = 0 Then
                            lng������λ���� = 0
                        Else
                            Do While Not rsUnit.EOF
                                lng������λ���� = lng������λ���� + Int(Val(Nvl(rsUnit!����)) * Val(Nvl(rsTemp!��Լ��)) / 100)
                                rsUnit.MoveNext
                            Loop
                        End If
                    End If
                End If
                If Not IsNull(rsTemp!��Լ��) Then
                    If Val(Nvl(rsTemp!��Լ��)) + lng������λ���� - lng������λ�������� >= Val(Nvl(rsTemp!��Լ��)) Then
                        MsgBox "��ǰԤԼ���볬������������" & Val(Nvl(rsTemp!��Լ��)) & "(���а���������λ��������" & lng������λ���� & "),���ܼ���ԤԼ!", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    CheckLimit = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SaveData() As Boolean
    On Error GoTo errHandle
    Dim i As Integer, k As Integer, int�۸񸸺� As Integer, strDay As String, j As Integer
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim str�Ǽ�ʱ�� As String, lngSN As Long, str����ʱ�� As String, str���ʽ As String
    Dim lng�Һſ���ID As Long, byt���� As Byte, strNO As String, cllPro As New Collection, rsCheck As ADODB.Recordset
    Dim strҽ�� As String, lngҽ��ID As Long, strSQL As String, blnNoDoc As Boolean, dat����ʱ�� As Date
    Dim bytMode As Byte, datԤԼʱ�� As Date
    Dim strResult As String, blnר�Һ� As Boolean
    
    If vsfPlan.RowData(vsfPlan.Row) = "" Then
        MsgBox "��ѡ��һ�����Ž���ԤԼ!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If vsfList.Visible Then
        If vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col) = "" Then
            MsgBox "��ѡ��һ����Ч����Ž���ԤԼ!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If CheckLimit(Val(vsfPlan.RowData(vsfPlan.Row))) = False Then Exit Function
    
    If Not mrsInfo Is Nothing Then
        strSQL = "Select Zl_Fun_���˹Һż�¼_Check([1],[2],[3],[4],[5],[6]) As ����� From Dual"
        bytMode = 1
        datԤԼʱ�� = dtpMain.Selection.Blocks(0).DateBegin
        
        blnר�Һ� = vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("ҽ��")) <> ""
        Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytMode, Val(Nvl(mrsInfo!����ID)), vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�ű�")), _
                                                Val(vsfPlan.RowData(vsfPlan.Row)), datԤԼʱ��, IIf(blnר�Һ�, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!�����)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "��Ч�Լ��ʧ��,�޷�������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    ReadRegistPrice Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ĿID"))), False, False, txtFeeType.Text, rsItems, rsIncomes
    str�Ǽ�ʱ�� = "To_Date('" & zlDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss')"
    strDay = zlGet��ǰ���ڼ�(dtpMain.Selection.Blocks(0).DateBegin)
    If vsfList.Visible Then
        lngSN = vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)
    End If
    If lngSN <> 0 Then
        If MsgBox("�Ƿ�ԤԼ���(" & lngSN & ")?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Function
    Else
        If MsgBox("�Ƿ�ԤԼ�ú�?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Function
    End If
    
    If lngSN <> 0 Then
        If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʱ��"))) = 1 Then
            strSQL = "Select ��ʼʱ�� From �ٴ�������ſ��� Where ��¼ID=[1] And ���=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), lngSN)
            If Not rsTemp.EOF Then
                dat����ʱ�� = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd hh:mm:ss"))
                str����ʱ�� = "To_Date('" & Format(rsTemp!��ʼʱ��, "yyyy-mm-dd hh:mm:ss") & " ','YYYY-MM-DD HH24:MI:SS')"
            Else
                dat����ʱ�� = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                str����ʱ�� = "To_Date('" & Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00") & " ','YYYY-MM-DD HH24:MI:SS')"
            End If
        Else
            dat����ʱ�� = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
            str����ʱ�� = "To_Date('" & Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00") & " ','YYYY-MM-DD HH24:MI:SS')"
        End If
    Else
        dat����ʱ�� = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
        str����ʱ�� = "To_Date('" & Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00") & " ','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    If dat����ʱ�� < DateAdd("n", -1 * mlngԤԼ��Чʱ��, zlDatabase.Currentdate) Then
        MsgBox "ԤԼʱ��С���˿�ԤԼʱ��(" & Format(DateAdd("n", -1 * mlngԤԼ��Чʱ��, zlDatabase.Currentdate), "hh:mm:ss") & "),�޷�ԤԼ!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʱ��"))) = 1 And Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ſ���"))) = 1) Then
        If Check��Чʱ���(Val(vsfPlan.RowData(vsfPlan.Row)), dat����ʱ��) = False Then
            MsgBox "��ǰѡ��ĳ����¼��" & Format(dat����ʱ��, "yyyy-mm-dd hh:mm:ss") & "������,������Һ�ʱ��!", vbInformation, gstrSysName
            If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
            Exit Function
        End If
    End If
    
    If lngSN = 0 Then
        strSQL = "Select Zl_Fun_Get�ٴ�����ԤԼ״̬([1],[2]) As ԤԼ��� From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), dat����ʱ��)
    Else
        strSQL = "Select Zl_Fun_Get�ٴ�����ԤԼ״̬([1],[2],[3]) As ԤԼ��� From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), dat����ʱ��, lngSN)
    End If
    If rsTemp.EOF Then
        MsgBox "��ǰѡ��ĳ����¼�޷�ԤԼ!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!ԤԼ���), 1, 1)) <> 0 Then
            MsgBox "��ǰѡ��ĳ����¼�޷�ԤԼ!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!ԤԼ���), InStr(Nvl(rsTemp!ԤԼ���), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select Zl_�ٴ���������_Check([1],[2],[3]) As �����Լ�� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), txtGender.Text, txtAge.Text)
'    If rsTemp.EOF Then
'        MsgBox "��ǰѡ��Ĳ��˲����øúű�!", vbInformation, gstrSysName
'        Exit Function
    If Not rsTemp.EOF Then
        If Val(Mid(Nvl(rsTemp!�����Լ��), 1, 1)) <> 0 Then
            MsgBox "��ǰѡ��Ĳ��˲����øúű�!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!�����Լ��), InStr(Nvl(rsTemp!�����Լ��), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng�Һſ���ID = Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ID")))
    byt���� = IIf(Check����(Val(txtName.Tag), lng�Һſ���ID), 1, 0)
    strNO = zlDatabase.GetNextNo(12)
    
    With vsfPlan
        If .TextMatrix(.Row, .ColIndex("����ҽ������")) <> "" Then
            If dat����ʱ�� >= CDate(.TextMatrix(.Row, .ColIndex("���￪ʼʱ��"))) And dat����ʱ�� <= CDate(.TextMatrix(.Row, .ColIndex("������ֹʱ��"))) Then
                strҽ�� = .TextMatrix(.Row, .ColIndex("����ҽ������"))
                lngҽ��ID = .TextMatrix(.Row, .ColIndex("����ҽ��ID"))
            Else
                strҽ�� = .TextMatrix(.Row, .ColIndex("ҽ��"))
                lngҽ��ID = Val(.TextMatrix(.Row, .ColIndex("ҽ��ID")))
            End If
        End If
    End With
    
    
    strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ���� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, Nvl(mrsInfo!ҽ�Ƹ��ʽ))
    If rsTemp.RecordCount <> 0 Then
        str���ʽ = Nvl(rsTemp!����)
    Else
        strSQL = "Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
        If rsTemp.RecordCount <> 0 Then
            str���ʽ = Nvl(rsTemp!����)
        End If
    End If
    
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int�۸񸸺� = k
        rsIncomes.Filter = "��ĿID=" & rsItems!��ĿID
        For j = 1 To rsIncomes.RecordCount
            strSQL = _
            "zl_���˹Һż�¼_����_INSERT(" & ZVal(vsfPlan.RowData(vsfPlan.Row)) & "," & Val(Nvl(mrsInfo!����ID)) & "," & IIf(IsNull(mrsInfo!�����), "NULL", mrsInfo!�����) & ",'" & txtName.Text & "','" & txtGender.Text & "'," & _
                     "'" & txtAge.Text & "','" & str���ʽ & "','" & txtFeeType.Text & "','" & strNO & "'," & _
                     "'" & "" & "'," & k & "," & IIf(int�۸񸸺� = k, "NULL", int�۸񸸺�) & "," & IIf(rsItems!���� = 2, 1, "NULL") & "," & _
                     "'" & rsItems!��� & "'," & rsItems!��ĿID & "," & rsItems!���� & "," & rsIncomes!���� & "," & _
                     rsIncomes!������ĿID & ",'" & rsIncomes!�վݷ�Ŀ & "','" & "" & "'," & _
                      rsIncomes!Ӧ�� & "," & rsIncomes!ʵ�� & "," & _
                     lng�Һſ���ID & "," & UserInfo.����ID & "," & IIf(rsItems!ִ�п���ID = 0, lng�Һſ���ID, rsItems!ִ�п���ID) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                     str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                     "'" & strҽ�� & "'," & ZVal(lngҽ��ID) & "," & IIf(rsItems!���� = 3, 1, IIf(rsItems!���� = 4, 2, 0)) & "," & Val(Nvl(mrsInfo!��Ŀ����)) & "," & _
                     "'" & vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�ű�")) & "','" & IIf(strҽ�� = UserInfo.����, "", "") & "'," & ZVal(0) & "," & "NULL" & "," & _
                     ZVal(IIf(k = 1, 0, 0)) & "," & ZVal(IIf(k = 1, 0, 0)) & "," & _
                     ZVal(IIf(k = 1, 0, 0)) & "," & ZVal(Nvl(rsItems!���մ���id, 0)) & "," & _
                     ZVal(Nvl(rsItems!������Ŀ��, 0)) & "," & ZVal(Nvl(rsIncomes!ͳ����, 0)) & "," & _
                     "'" & "" & "'," & 1 & "," & 0 & ",'" & rsItems!���ձ��� & "'," & byt���� & "," & ZVal(lngSN) & ",Null," & _
                     1 & ",'" & "" & "'," & _
                     0 & ","
            '�����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSQL = strSQL & "NULL" & ","
            '���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
            strSQL = strSQL & "NULL" & ","
            '����_In       ����Ԥ����¼.����%Type := Null,
            strSQL = strSQL & "'" & "" & "',"
            '������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSQL = strSQL & " NULL,"
            '����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            strSQL = strSQL & " NULL,"
            '������λ_In   ����Ԥ����¼.������λ%Type := Null
            strSQL = strSQL & " NULL,"
            '  ��������_In   Number:=0
            strSQL = strSQL & 0 & ","
            '  ����_IN       ���˹Һż�¼.����%type:=null,
            strSQL = strSQL & "NULL" & ","
            '  ����ģʽ_IN   NUMBER :=0,
            strSQL = strSQL & 0 & ","
            '  ���ʷ���_IN Number:=0
            strSQL = strSQL & 0 & ","
            '  �˺�����_IN Number:=1
            strSQL = strSQL & 0 & ","
            '  ��Ԥ������ids_In Varchar2 := Null
            strSQL = strSQL & "'" & "" & "',"
            '  �������˷ѱ�_In Number := 0
            strSQL = strSQL & 0 & ")"
            
            Call zlAddArray(cllPro, strSQL)
            '����:31187:���ҺŻ��ܵ�������
            If Val(vsfPlan.RowData(vsfPlan.Row)) <> 0 And k = 1 Then
                If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("ҽ��")) = "" Then blnNoDoc = True
                strSQL = "zl_���˹ҺŻ���_Update("
                '  ҽ������_In   �ҺŰ���.ҽ������%Type,
                strSQL = strSQL & IIf(blnNoDoc, "Null,", "'" & strҽ�� & "',")
                '  ҽ��id_In     �ҺŰ���.ҽ��id%Type,
                strSQL = strSQL & "" & IIf(blnNoDoc, "0,", ZVal(lngҽ��ID) & ",")
                '  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
                strSQL = strSQL & "" & Val(Nvl(rsItems!��ĿID)) & ","
                '  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
                strSQL = strSQL & "" & IIf(Val(Nvl(rsItems!ִ�п���ID)) = 0, lng�Һſ���ID, Val(Nvl(rsItems!ִ�п���ID))) & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                strSQL = strSQL & "" & str����ʱ�� & ","
                '  ԤԼ��־_In   Number := 0  --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����,3-�շ�ԤԼ
                strSQL = strSQL & 1 & ","
                '  ����_In       �ҺŰ���.����%Type := Null
                strSQL = strSQL & "'" & vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("�ű�")) & "',0,"
                strSQL = strSQL & "" & Val(vsfPlan.RowData(vsfPlan.Row)) & ")"
                Call zlAddArray(cllPro, strSQL)
            End If
            
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, False, False
    
    strSQL = "Select ID From ���˹Һż�¼ Where No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    strSQL = "Zl_���߷�������_����("
    strSQL = strSQL & mlng��ϢID & ","
    strSQL = strSQL & "Null,'"
    strSQL = strSQL & UserInfo.���� & "','"
    strSQL = strSQL & UserInfo.��� & "',"
    strSQL = strSQL & Val(rsTemp!ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Call mfrmMain.RefreshData
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Check��Чʱ���(lng��¼ID As Long, datTime As Date) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And [2] Between ��ʼʱ�� And ��ֹʱ�� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, datTime)
    
    If rsTemp.EOF Then
        Check��Чʱ��� = False
    Else
        strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And [2] Between Nvl(ͣ�￪ʼʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(ͣ����ֹʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, datTime)
        If rsTemp.EOF Then
            Check��Чʱ��� = True
        Else
            Check��Чʱ��� = False
        End If
    End If
End Function

Private Sub cboTime_Click()
    If mblnNotClick Then Exit Sub
    Call ShowRow
End Sub

Private Sub cboTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdDirectApp_Click()
    Call SaveData
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Public Sub LoadData(frmMain As Object, ByVal lngID As Long)
    On Error GoTo errHandle
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim datBegin As Date, datEnd As Date
    Dim datNow As Date
    Set mfrmMain = frmMain
    Call ClearData
    mlng��ϢID = lngID
    strSQL = "Select a.����Id, b.�����, b.����, b.���֤��, b.ҽ�Ƹ��ʽ, d.��Ŀ����, b.�Ա�, b.����, b.��������, b.�ѱ�, b.��ͥ�绰, b.��ͥ��ַ, b.����, b.����, b.ְҵ, b.����״��, a.֪ͨԭ�� As �Ǽ�ԭ��, a.�Ǽ���, a.��ʼʱ��," & vbNewLine & _
            "       a.��ֹʱ��, c.���� As ԤԼ����, a.ҽ������ As ԤԼҽ��, a.��ĿID, d.���� As ��Ŀ����, a.�Ǽ�ʱ�� " & vbNewLine & _
            "From ���˷�����Ϣ��¼ A, ������Ϣ B, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where a.Id = [1] And a.����id = b.����id And a.����id = c.Id And a.��ĿId = d.Id "
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If mrsInfo.EOF Then
        MsgBox "��ȡ������Ϣʧ��,�޷����������Ϣ!"
        Exit Sub
    End If
    
    '�۸�ȼ�
    If gintPriceGradeStartType >= 2 Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), 0, Nvl(mrsInfo!ҽ�Ƹ��ʽ, ""), , , mstrPriceGrade)
    Else
        mstrPriceGrade = gstrPriceGrade
    End If
    
    txtName.Text = Nvl(mrsInfo!����)
    txtName.Tag = Nvl(mrsInfo!����ID)
    txtID.Text = Nvl(mrsInfo!���֤��)
    txtGender.Text = Nvl(mrsInfo!�Ա�)
    txtAge.Text = Nvl(mrsInfo!����)
    txtBirth.Text = Nvl(mrsInfo!��������)
    txtFeeType.Text = Nvl(mrsInfo!�ѱ�)
    txtPhone.Text = Nvl(mrsInfo!��ͥ�绰)
    txtAddress.Text = Nvl(mrsInfo!��ͥ��ַ)
    txtNation.Text = Nvl(mrsInfo!����)
    txtRace.Text = Nvl(mrsInfo!����)
    txtJob.Text = Nvl(mrsInfo!ְҵ)
    txtMarriage.Text = Nvl(mrsInfo!����״��)
    txtState.Text = Nvl(mrsInfo!�Ǽ�ԭ��)
    txtReger.Text = Nvl(mrsInfo!�Ǽ���)
    txtTimeBegin.Text = Format(Nvl(mrsInfo!��ʼʱ��), "yyyy-mm-dd")
    txtTimeEnd.Text = Format(Nvl(mrsInfo!��ֹʱ��), "yyyy-mm-dd")
    txtDept.Text = Nvl(mrsInfo!ԤԼ����)
    txtDoc.Text = Nvl(mrsInfo!ԤԼҽ��)
    txtItem.Text = Nvl(mrsInfo!��Ŀ����)
    txtMoney.Text = Format(Get��Ŀ���(Val(Nvl(mrsInfo!��ĿID)), mstrPriceGrade), "0.00")
    txtRegTime.Text = Format(Nvl(mrsInfo!�Ǽ�ʱ��), "yyyy-mm-dd hh:mm:ss")
    mblnInit = True
    mblnNotClick = True
    txtFilter.Text = txtDoc.Text
    mblnNotClick = False
    datBegin = CDate(txtTimeBegin.Text & " 00:00:00")
    datEnd = CDate(txtTimeEnd.Text & " 23:59:59")
    datNow = zlDatabase.Currentdate
    strSQL = "Select �������� " & vbNewLine & _
            "From �ٴ������¼" & vbNewLine & _
            "Where ����ҽ������ Is Null And ҽ������ = [1] And �������� Between [2] And [3] Order By ��������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtDoc.Text, datBegin, datEnd)
    mblnNotClick = True
    If rsTemp.EOF Then
        If datNow > datBegin Then
            dtpMain.SelectRange datNow, datNow
            dtpMain.Select datNow
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        Else
            dtpMain.SelectRange datBegin, datBegin
            dtpMain.Select datBegin
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        End If
    Else
        If datNow > CDate(rsTemp!��������) Then
            dtpMain.SelectRange datNow, datNow
            dtpMain.Select datNow
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        Else
            dtpMain.SelectRange CDate(rsTemp!��������), CDate(rsTemp!��������)
            dtpMain.Select CDate(rsTemp!��������)
            dtpMain.EnsureVisibleSelection
            dtpMain.RedrawControl
        End If
    End If
    Do While Not rsTemp.EOF
        '����Ⱦɫ
        rsTemp.MoveNext
    Loop
    mblnNotClick = False
    
    Call LoadPlan
    Call ShowRow
    mblnInit = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPlan()
    Dim strSQL As String, rsPlan As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim datApp As Date, i As Integer, dblMoney As Double, blnAdd As Boolean
    Dim strTime() As String, lngLeft As Long, intChar As Integer
    Dim str��Ŀids As String
    On Error GoTo errH
    datApp = dtpMain.Selection.Blocks(0).DateBegin
    mdatCache = datApp
    strSQL = "Select a.id, b.����, b.���� As �ű�, c.���� As ����, c.���� As ���Ҽ���, b.����Id, a.�ϰ�ʱ�� As ʱ��, " & _
            "        d.���� As ��Ŀ, zlSpellcode(d.����) As ��Ŀ����, a.����ҽ��ID, a.����ҽ������, a.ҽ��Id, a.ҽ������ As ҽ��, e.���� As ҽ������, " & _
            "        a.��Ŀid, a.��Լ�� As ��Լ, a.��Լ�� As ��Լ, Nvl(a.�Ƿ��ʱ��,0) As ��ʱ��, Nvl(a.�Ƿ���ſ���,0) As ��ſ���, " & _
            "        a.��������, a.ȱʡԤԼʱ��, a.���￪ʼʱ��, a.������ֹʱ��, a.��ʼʱ��, a.��ֹʱ�� " & vbNewLine & _
            "From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D, ��Ա�� E" & vbNewLine & _
            "Where a.��Դid = b.Id  And Nvl(C.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.��Ŀid = d.Id And b.����id = c.Id And (c.վ�� Is Null Or c.վ�� = '" & gstrNodeNo & "') " & _
            "      And (a.�������� = [1] Or a.�������� = [2]) And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) Or Exists (Select 1 From �ٴ�������ſ��� C,�ٴ������¼ D Where D.ID=A.ID And C.��¼ID=D.ID And Nvl(C.�Ƿ�ͣ��,0) = 0 And D.�Ƿ���ſ��� =1 And D.�Ƿ��ʱ�� = 1 And C.��ʼʱ�� <> C.��ֹʱ��)) " & _
            "      And Nvl(a.�Ƿ񷢲�,0)=1 And a.ҽ��Id = e.Id(+) And Nvl(a.ԤԼ����,0) <> 1 " & _
            "      And Not Exists (Select 1 From �ٴ������¼ Where Id=a.Id And ��ֹʱ�� < [3]) And a.��ʼʱ�� >= [4] And Sysdate + zl_Fun_GetAppointmentDays + Decode(Nvl(B.ԤԼ����," & gintԤԼ���� & "),0,15,Nvl(B.ԤԼ����," & gintԤԼ���� & ")" & ") >= [1] " & _
            "      And [3] Not Between Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) And Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) "
    If Format(datApp, "yyyy-mm-dd") = Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
        Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datApp, datApp - 1, zlDatabase.Currentdate, gdatRegistTime)
    Else
        Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datApp, datApp - 1, datApp, gdatRegistTime)
    End If
    mstrʱ��s = ""
    mblnNotClick = True
    cboTime.Clear
    cboTime.AddItem "����"
    mblnNotClick = False
    vsfPlan.Redraw = flexRDNone
    vsfPlan.Rows = 1
    vsfPlan.Clear 1
    vsfPlan.Rows = 2
    If rsPlan.RecordCount <> 0 Then
        str��Ŀids = ""
        Do While Not rsPlan.EOF
            If InStr("," & str��Ŀids & ",", "," & Val(Nvl(rsPlan!��ĿID)) & ",") = 0 Then
                str��Ŀids = str��Ŀids & "," & Val(Nvl(rsPlan!��ĿID))
            End If
            rsPlan.MoveNext
        Loop
        
        rsPlan.MoveFirst
    End If
    
    If str��Ŀids <> "" Then
        str��Ŀids = Mid(str��Ŀids, 2)
    End If
    Set rsTemp = Get��Ŀ��Ϣ(str��Ŀids, mstrPriceGrade)
    
    Do While Not rsPlan.EOF
        blnAdd = True
        
        If blnAdd Then
            With vsfPlan
                rsTemp.Filter = "��ĿID=" & Val(Nvl(rsPlan!��ĿID))
                .RowData(.Rows - 1) = Val(Nvl(rsPlan!ID))
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsPlan!����)
                .TextMatrix(.Rows - 1, .ColIndex("�ű�")) = Nvl(rsPlan!�ű�)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsPlan!����)
                .TextMatrix(.Rows - 1, .ColIndex("ʱ��")) = Nvl(rsPlan!ʱ��)
                If InStr("," & mstrʱ��s & ",", "," & Nvl(rsPlan!ʱ��) & ",") = 0 Then
                    mstrʱ��s = mstrʱ��s & "," & Nvl(rsPlan!ʱ��)
                End If
                .TextMatrix(.Rows - 1, .ColIndex("ҽ��")) = Nvl(rsPlan!ҽ��)
                If Nvl(rsPlan!����ҽ������) <> "" Then
                    .Cell(flexcpData, .Rows - 1, .ColIndex("����ҽ��")) = Nvl(rsPlan!����ҽ������) & "(" & Format(Nvl(rsPlan!���￪ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsPlan!������ֹʱ��), "hh:mm") & ")"
                    .TextMatrix(.Rows - 1, .ColIndex("����ҽ��")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("����ҽ������")) = Nvl(rsPlan!����ҽ������)
                    .TextMatrix(.Rows - 1, .ColIndex("����ҽ��ID")) = Nvl(rsPlan!����ҽ��id)
                    .TextMatrix(.Rows - 1, .ColIndex("���￪ʼʱ��")) = Nvl(rsPlan!���￪ʼʱ��)
                    .TextMatrix(.Rows - 1, .ColIndex("������ֹʱ��")) = Nvl(rsPlan!������ֹʱ��)
                End If
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Nvl(rsPlan!��Ŀ)
                If rsTemp.EOF Then
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = "0.00"
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(Val(Nvl(rsTemp!���)), "0.00")
                End If
                .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(Get��Ŀ���(Val(Nvl(rsPlan!��ĿID)), mstrPriceGrade), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("��Լ")) = Nvl(rsPlan!��Լ)
                .TextMatrix(.Rows - 1, .ColIndex("��Լ")) = Val(Nvl(rsPlan!��Լ))
                .TextMatrix(.Rows - 1, .ColIndex("��ʱ��")) = Nvl(rsPlan!��ʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��ſ���")) = Nvl(rsPlan!��ſ���)
                .TextMatrix(.Rows - 1, .ColIndex("��ĿID")) = Nvl(rsPlan!��ĿID)
                .TextMatrix(.Rows - 1, .ColIndex("����ID")) = Nvl(rsPlan!����ID)
                .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = Nvl(rsPlan!ҽ��ID)
                .TextMatrix(.Rows - 1, .ColIndex("ҽ������")) = Nvl(rsPlan!ҽ������)
                .TextMatrix(.Rows - 1, .ColIndex("���Ҽ���")) = Nvl(rsPlan!���Ҽ���)
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = Nvl(rsPlan!��Ŀ����)
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(Nvl(rsPlan!��������), "yyyy-mm-dd")
                .TextMatrix(.Rows - 1, .ColIndex("ԤԼʱ��")) = Format(Nvl(rsPlan!ȱʡԤԼʱ��), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = Format(Nvl(rsPlan!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = Format(Nvl(rsPlan!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
                
                .Rows = .Rows + 1
            End With
        End If
        rsPlan.MoveNext
    Loop
    mblnNotClick = True
    If mstrʱ��s <> "" Then
        mstrʱ��s = Mid(mstrʱ��s, 2)
        strTime = Split(mstrʱ��s, ",")
        For i = 0 To UBound(strTime)
            cboTime.AddItem strTime(i)
        Next i
    End If
    cboTime.ListIndex = 0
    mblnNotClick = False

    If rsPlan.RecordCount = 0 Then
        mfrmMain.ShowPanelText "��ǰ����û�п���ԤԼ�ĺ���,�޷�ԤԼ!"
        vsfPlan.Redraw = flexRDDirect
        vsfList.Visible = False
        vsfPlan.Height = picReg.ScaleHeight - 500
        vsfPlan.Select 1, 1
        mblnUnload = True
        Exit Sub
    Else
        mfrmMain.ShowPanelText ""
    End If
    Call ShowRow
    If vsfPlan.Rows <> 2 Then vsfPlan.Rows = vsfPlan.Rows - 1
    vsfPlan.Select 1, 1
    vsfPlan.AutoSize 0, vsfPlan.Cols - 1
    zl_vsGrid_Para_Restore 1115, vsfPlan, Me.Name, "vsfPlan"
    vsfPlan.Redraw = flexRDDirect
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowRow()
    Dim i As Integer, blnHide As Boolean
    Dim blnEnable As Boolean, strTimeRange As String
    On Error GoTo errH
    If vsfPlan.Rows = 2 And vsfPlan.TextMatrix(1, vsfPlan.ColIndex("�ű�")) = "" Then Exit Sub
    If cboTime.Text <> "����" Then strTimeRange = cboTime.Text
    With vsfPlan
        For i = 1 To .Rows - 1
            blnHide = False
            If txtFilter <> "" Then
                blnHide = True
                If .TextMatrix(i, .ColIndex("�ű�")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("����")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("����")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("��Ŀ")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("ҽ��")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("���Ҽ���"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("ҽ������"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("��Ŀ����"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
            End If
            If strTimeRange <> .TextMatrix(i, .ColIndex("ʱ��")) And strTimeRange <> "" Then blnHide = True
'            If InStr(strTimeRange & ",", .TextMatrix(i, .ColIndex("ʱ��"))) > 0 Then blnHide = True
            .RowHidden(i) = blnHide
        Next i
    End With
    blnEnable = False
    With vsfPlan
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                blnEnable = True
                .Select i, 1
                Call vsfPlan_EnterCell
                Exit For
            End If
        Next i
    End With
    If blnEnable = False Then
        If mblnKeyPress Then
            vsfList.Visible = False
            vsfPlan.Height = picReg.ScaleHeight - 500
            vsfPlan.Select 1, 1
        Else
            txtFilter.Text = ""
        End If
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearData()
    On Error GoTo errHandle
    txtName.Text = ""
    txtID.Text = ""
    txtGender.Text = ""
    txtAge.Text = ""
    txtBirth.Text = ""
    txtFeeType.Text = ""
    txtPhone.Text = ""
    txtAddress.Text = ""
    txtNation.Text = ""
    txtRace.Text = ""
    txtJob.Text = ""
    txtMarriage.Text = ""
    txtState.Text = ""
    txtReger.Text = ""
    txtTimeBegin.Text = ""
    txtTimeEnd.Text = ""
    txtDept.Text = ""
    txtDoc.Text = ""
    txtItem.Text = ""
    txtMoney.Text = ""
    cboTime.Clear
    cboTime.AddItem "����"
    cboTime.ListIndex = cboTime.NewIndex
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDate_Change()
    Dim str���� As String, i As Integer, lngRow As Long
    Dim str����ʱ�� As String
    If Not dtpMain.Visible Then Exit Sub
    If Not dtpMain.Enabled Then Exit Sub
    
    str���� = Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-MM-dd")

    If str���� = "" Then str���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    str����ʱ�� = str���� & " " & Format(dtpDate.Value, "hh:mm:00")
    lngRow = 0
    If CDate(str����ʱ��) > CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ֹʱ��"))) Then
        '����ʱ��İ��ţ�����Ѱ�Ҷ�λ
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("�ű�")) = .TextMatrix(i, .ColIndex("�ű�")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("��ֹʱ��"))) >= CDate(str����ʱ��) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    ElseIf CDate(str����ʱ��) < CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʼʱ��"))) Then
        '����ʱ��İ��ţ�����Ѱ�Ҷ�λ
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("�ű�")) = .TextMatrix(i, .ColIndex("�ű�")) And _
                    CDate(vsfPlan.TextMatrix(i, vsfPlan.ColIndex("��ʼʱ��"))) <= CDate(str����ʱ��) Then
                    lngRow = i
                    Exit For
                End If
            End With
        Next i
    End If
    If lngRow <> 0 Then
        mblnAppointmentChange = True
        vsfPlan.Select lngRow, 1
        mblnAppointmentChange = False
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call dtpDate_Change
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call InitPanel
    Call InitGrid
    Call InitPara
End Sub

Private Sub InitPara()
    mlngԤԼ��Чʱ�� = -1 * Val(Split(zlDatabase.GetPara("ԤԼ����ʱ��", glngSys, 1111, "1|60") & "|", "|")(1))
    mintר�Һ�ԤԼ���� = Val(zlDatabase.GetPara("ר�Һ�ԤԼ����", glngSys, , 0))
    mint����ԤԼ������ = Val(zlDatabase.GetPara("����ԤԼ������", glngSys, 1111, 0))
    mintͬ����Լ�� = Val(zlDatabase.GetPara("����ͬ����ԼN����", glngSys, 1111, 0))
End Sub

Private Sub InitGrid()
    Dim i As Integer
    With vsfPlan
        .Cols = 26
        .Rows = 2
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "�ű�"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "ʱ��"
        .TextMatrix(0, 4) = "��Ŀ"
        .TextMatrix(0, 5) = "ҽ��"
        .TextMatrix(0, 6) = "����ҽ��"
        .TextMatrix(0, 7) = "���"
        .TextMatrix(0, 8) = "��Լ"
        .TextMatrix(0, 9) = "��Լ"
        .TextMatrix(0, 10) = "��ʱ��"
        .TextMatrix(0, 11) = "��ſ���"
        .TextMatrix(0, 12) = "��ĿID"
        .TextMatrix(0, 13) = "����ID"
        .TextMatrix(0, 14) = "ҽ��ID"
        .TextMatrix(0, 15) = "���Ҽ���"
        .TextMatrix(0, 16) = "ҽ������"
        .TextMatrix(0, 17) = "��Ŀ����"
        .TextMatrix(0, 18) = "��������"
        .TextMatrix(0, 19) = "ԤԼʱ��"
        .TextMatrix(0, 20) = "����ҽ������"
        .TextMatrix(0, 21) = "����ҽ��ID"
        .TextMatrix(0, 22) = "���￪ʼʱ��"
        .TextMatrix(0, 23) = "������ֹʱ��"
        .TextMatrix(0, 24) = "��ʼʱ��"
        .TextMatrix(0, 25) = "��ֹʱ��"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "��ʱ��" Or .ColKey(i) = "��ſ���" Or .ColKey(i) = "��ĿID" Or _
                .ColKey(i) = "����ID" Or .ColKey(i) = "ҽ��ID" Or .ColKey(i) = "���Ҽ���" _
                Or .ColKey(i) = "ҽ������" Or .ColKey(i) = "��Ŀ����" Or .ColKey(i) = "ԤԼʱ��" _
                Or .ColKey(i) = "����ҽ������" Or .ColKey(i) = "����ҽ��ID" Or .ColKey(i) = "���￪ʼʱ��" _
                Or .ColKey(i) = "������ֹʱ��" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Then .ColHidden(i) = True
        Next i
    End With
    With vsfList
        For i = 0 To .Rows - 1
            .RowHeight(i) = 500
            .Cell(flexcpFontBold, i, 0) = True
            .Cell(flexcpFontSize, i, 0) = 16
        Next i
    End With
    vsfList.Visible = False
    vsfPlan.Height = picReg.ScaleHeight - 500
End Sub

Private Sub dtpMain_SelectionChanged()
    Dim datNow As Date
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd") > Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
        If Format(datNow, "yyyy-mm-dd") > Format(mdatCache, "yyyy-mm-dd") Then
            mdatCache = datNow
        End If
        dtpMain.SelectRange mdatCache, mdatCache
        dtpMain.Select mdatCache
        dtpMain.EnsureVisibleSelection
        dtpMain.RedrawControl
    End If
    If mblnNotClick Then Exit Sub
    Call LoadPlan
    Call ShowRow
End Sub


Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save 1115, vsfPlan, Me.Name, "vsfPlan"
End Sub

Private Sub picDate_Resize()
    With dtpMain
'        .Height = picDate.ScaleHeight
        .Width = picDate.ScaleWidth
    End With
End Sub

Private Sub picReg_Resize()
    On Error Resume Next
    With vsfPlan
        .Width = picReg.ScaleWidth - 150
        If vsfList.Visible Then
            .Height = vsfList.Top - .Top - 100
            picSplit.Top = vsfList.Top - 60
        Else
            .Height = picReg.ScaleHeight - 500
        End If
    End With
    picSplit.Width = picReg.ScaleWidth
    With vsfList
        .Width = picReg.ScaleWidth - 150
        .Height = picReg.ScaleHeight - vsfList.Top
    End With
End Sub

Private Sub txtFilter_Change()
    If mblnNotClick Then Exit Sub
    mblnKeyPress = True
    Call ShowRow
    mblnKeyPress = False
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfList.TextMatrix(Row, Col) = "" Then Cancel = True
    If Not ((vsfList.Cell(flexcpForeColor, Row, Col) = vbBlack Or vsfList.Cell(flexcpForeColor, Row, Col) = 2) And vsfList.Cell(flexcpFontStrikethru, Row, Col) = False) Then Cancel = True
    vsfList.ComboList = "..."
    vsfList.CellButtonPicture = imgApp
End Sub


Private Sub vsfList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call SaveData
End Sub

Private Sub vsfPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer, blnMark As Boolean
    With vsfPlan
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                If blnMark Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HEEEEEE
                    blnMark = False
                Else
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80000005
                    blnMark = True
                End If
                If i = .Row Then
                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = 16772055
                End If
            End If
        Next i
'        If OldRow < .Rows Then
'            If OldRow Mod 2 = 1 Then
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
'            Else
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
'            End If
'        End If
'        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

Private Sub vsfPlan_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer, blnMark As Boolean
    With vsfPlan
'        For i = 1 To .Rows - 1
'            If .RowHidden(i) = False Then
'                If blnMark Then
'                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HEEEEEE
'                    blnMark = False
'                Else
'                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80000005
'                    blnMark = True
'                End If
'                If i = .Row Then
'                    .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = 16772055
'                End If
'            End If
'        Next i
'        If OldRow < .Rows Then
'            If OldRow Mod 2 = 1 Then
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
'            Else
'                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
'            End If
'        End If
'        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

Private Sub vsfPlan_EnterCell()
    Dim i As Integer, j As Integer, datApp As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset
    With vsfPlan
        If Val(.TextMatrix(.Row, .ColIndex("��ʱ��"))) = 1 Then
            dtpDate.Enabled = False
            With vsfList
                For i = 0 To .Rows - 1
                    .RowHeight(i) = 500
                    .Cell(flexcpFontBold, i, 0) = True
                    .Cell(flexcpFontSize, i, 0) = 16
                Next i
            End With
            vsfList.Visible = True
            picSplit.Visible = True
            vsfPlan.Height = vsfList.Top - .Top - 100
            picSplit.Top = vsfList.Top - 60
            vsfList.Height = picReg.ScaleHeight - vsfList.Top
            Call LoadTimePlan
        Else
            dtpDate.Enabled = True
            If mblnAppointmentChange = False Then
                If vsfPlan.TextMatrix(.Row, .ColIndex("ԤԼʱ��")) <> "" And IsDate(vsfPlan.TextMatrix(.Row, .ColIndex("ԤԼʱ��"))) = True Then
                    If Format(vsfPlan.TextMatrix(.Row, .ColIndex("��������")), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                        dtpDate.Value = vsfPlan.TextMatrix(.Row, .ColIndex("��ֹʱ��"))
                    Else
                        dtpDate.Value = vsfPlan.TextMatrix(.Row, .ColIndex("ԤԼʱ��"))
                    End If
                Else
                    dtpDate.Value = zlDatabase.Currentdate
                End If
            End If

            If Val(vsfPlan.TextMatrix(.Row, .ColIndex("��ſ���"))) = 0 Then
                vsfList.Visible = False
                picSplit.Visible = False
                vsfPlan.Height = picReg.ScaleHeight - 500
            Else
                With vsfList
                    For i = 0 To .Rows - 1
                        .RowHeight(i) = 350
                        For j = 0 To .Cols - 1
                            .Cell(flexcpFontBold, i, j) = True
                            .Cell(flexcpFontSize, i, j) = 16
                        Next j
                    Next i
                End With
                vsfList.Visible = True
                picSplit.Visible = True
                vsfPlan.Height = vsfList.Top - .Top - 100
                picSplit.Top = vsfList.Top - 60
                vsfList.Height = picReg.ScaleHeight - vsfList.Top
                Call LoadSerialPlan
            End If
            If vsfList.Visible Then
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʱ��"))) = 1 Then
                    If Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) <> 0 Then
                        strSQL = "Select ��ʼʱ�� From �ٴ�������ſ��� Where ��¼ID=[1] And ���=[2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)))
                        If Not rsTemp.EOF Then
                            datApp = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd hh:mm:ss"))
                        Else
                            datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                        End If
                    Else
                        datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                    End If
                Else
                    datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
                End If
            Else
                datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpDate.Value, "hh:mm:00"))
            End If
            If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")) <> "" And vsfPlan.Row <> 0 Then
                If datApp >= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��"))) And datApp <= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("������ֹʱ��"))) Then
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
                Else
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
                End If
            End If
        End If
    End With
    cmdDirectApp.Visible = vsfList.Visible = False
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsfPlan.Height + Y < 500 Or vsfList.Height - Y < 500 Then Exit Sub
                
        picSplit.Top = picSplit.Top + Y
        vsfPlan.Height = vsfPlan.Height + Y
        vsfList.Top = vsfList.Top + Y
        vsfList.Height = vsfList.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnNotClick = False Then
        With vsfList
            If .TextMatrix(NewRow, NewCol) = "" Then Cancel = True
            If Not ((.Cell(flexcpForeColor, NewRow, NewCol) = vbBlack Or .Cell(flexcpForeColor, NewRow, NewCol) = 2) And .Cell(flexcpFontStrikethru, NewRow, NewCol) = False) Then Cancel = True
        End With
    End If
End Sub

Private Sub vsfList_EnterCell()
    If vsfList.Row >= vsfList.Rows Then Exit Sub
    If vsfList.Col >= vsfList.Cols Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), ":") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "-") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "ԤԼ") > 0 Then
        dtpDate.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(0), "-")(0)
    Else
        dtpDate.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(1), "-")(0)
    End If
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "��") > 0 Then
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
    Else
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
    End If
End Sub

Private Sub LoadSerialPlan()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intCurrentTime As Integer, intCol As Integer
    Dim blnFind As Boolean, i As Integer, j As Integer
    
    vsfList.Redraw = flexRDNone
    vsfList.Clear
    vsfList.Rows = 2
    vsfList.Cols = 10
    vsfList.FixedRows = 0
    vsfList.FixedCols = 0
    intCol = 0
    
    strSQL = "Select ���, ��ʼʱ��, ��ֹʱ��, �Ƿ�ԤԼ, �Һ�״̬, ����, ���� From �ٴ�������ſ��� Where ��¼id = [1] And Nvl(�Ƿ�ԤԼ,0)=1 Order By ���,��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
    Do While Not rsTemp.EOF
        With vsfList
            .TextMatrix(.Rows - 1, intCol) = Nvl(rsTemp!���)
            .Cell(flexcpData, .Rows - 1, intCol) = Nvl(rsTemp!���)
            Select Case Val(Nvl(rsTemp!�Һ�״̬))
                Case 0
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
                Case 1 '�ѹ�
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
                    .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                Case 2
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbGreen
                Case 3
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlue
                Case 4
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
                Case 5
                    .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
            End Select
            If Val(Nvl(rsTemp!�Ƿ�ԤԼ)) = 0 Then
                .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
            End If
            intCol = intCol + 1
            If intCol > 9 Then
                intCol = 0
                .Rows = .Rows + 1
            End If
        End With
        rsTemp.MoveNext
    Loop
    With vsfList
        For i = 0 To .Rows - 1
            .RowHeight(i) = 400
        Next i
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                .Cell(flexcpFontBold, i, j) = True
                .Cell(flexcpFontSize, i, j) = 10
            Next j
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 700
            .ColAlignment(i) = flexAlignCenterCenter
        Next i
    End With
    blnFind = False
    With vsfList
        For i = 0 To .Rows - 1
            If blnFind = False Then
                For j = 0 To .Cols - 1
                    If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False And vsfList.TextMatrix(i, j) <> "" Then
                        .Select i, j
                        Call vsfList_EnterCell
                        blnFind = True
                        Exit For
                    End If
                Next j
            End If
        Next i
        mblnNotClick = True
        If blnFind = False Then .Select 0, 0
        mblnNotClick = False
    End With
    vsfList.Rows = vsfList.Rows - 1
    vsfList.RowHidden(0) = True
    vsfList.Redraw = flexRDDirect
End Sub

Private Sub LoadTimePlan()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intCurrentTime As Integer, intCol As Integer
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean, rsUsed As ADODB.Recordset
    Dim rsUnit As ADODB.Recordset, lng������λ���� As Long, lng�ѹ����� As Long
    Dim datTime As Date
    Dim datNow As Date
    vsfList.Redraw = flexRDNone
    vsfList.Clear
    vsfList.Rows = 1
    vsfList.Cols = 2
    vsfList.FixedRows = 0
    vsfList.FixedCols = 1
    intCol = 0
    intCurrentTime = -1
    datTime = dtpMain.Selection.Blocks(0).DateBegin
    datNow = zlDatabase.Currentdate
    If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ſ���"))) = 1 Then
        strSQL = "Select ���, To_Char(��ʼʱ��,'hh24:mi:ss') As ��ʼʱ��, ��ʼʱ�� As ���ʱ��, To_Char(��ֹʱ��,'hh24:mi:ss') As ��ֹʱ��, �Ƿ�ԤԼ, Decode(�Ƿ�ͣ��,1,6,�Һ�״̬) As �Һ�״̬" & _
                " From �ٴ�������ſ��� Where ��¼id = [1] And Nvl(�Ƿ�ԤԼ,0)=1 And ��ʼʱ�� <> ��ֹʱ�� Order By ���,��ʼʱ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select ��� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id=[1] And ����=1 And ���Ʒ�ʽ=3"
        Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        Do While Not rsTemp.EOF
            rsUnit.Filter = "���=" & Val(rsTemp!���)
            If rsUnit.EOF Then
                lng������λ���� = 0
            Else
                lng������λ���� = 1
            End If
            With vsfList
                If intCurrentTime = -1 Then
                    intCurrentTime = Val(Split(Nvl(rsTemp!��ʼʱ��), ":")(0))
                    .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                    intCol = intCol + 1
                Else
                    If intCurrentTime = Val(Split(Nvl(rsTemp!��ʼʱ��), ":")(0)) Then
                        intCol = intCol + 1
                    Else
                        .Rows = .Rows + 1
                        intCurrentTime = Val(Split(Nvl(rsTemp!��ʼʱ��), ":")(0))
                        .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                        intCol = 1
                    End If
                End If
                
                If intCol >= .Cols Then .Cols = .Cols + 1
                If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) <> "" And _
                   Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")), "yyyy-mm-dd hh:mm:ss") And _
                   Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("������ֹʱ��")), "yyyy-mm-dd hh:mm:ss") Then
                    .TextMatrix(.Rows - 1, intCol) = Nvl(rsTemp!���) & "(��)" & vbCrLf & Format(Nvl(rsTemp!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "hh:mm")
                Else
                    .TextMatrix(.Rows - 1, intCol) = Nvl(rsTemp!���) & vbCrLf & Format(Nvl(rsTemp!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "hh:mm")
                End If
                .Cell(flexcpData, .Rows - 1, intCol) = Nvl(rsTemp!���)
                Select Case Val(Nvl(rsTemp!�Һ�״̬))
                    Case 0
                        If lng������λ���� = 0 Then
                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
                        Else
                            .Cell(flexcpForeColor, .Rows - 1, intCol) = &HFF00FF
                        End If
                    Case 1 '�ѹ�
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                    Case 2
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbGreen
                    Case 3
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlue
                    Case 4
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
                    Case 5
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                    Case 6
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End Select
                If Val(Nvl(rsTemp!�Ƿ�ԤԼ)) = 0 Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
                If CDate(Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlngԤԼ��Чʱ��, datNow) Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
                If Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
            End With
            rsTemp.MoveNext
        Loop
    Else
        strSQL = "Select ���, To_Char(��ʼʱ��,'hh24:mi:ss') As ��ʼʱ��, ��ʼʱ�� As ���ʱ��, To_Char(��ֹʱ��,'hh24:mi:ss') As ��ֹʱ��, ����, �Ƿ�ԤԼ From �ٴ�������ſ��� Where ��¼id = [1] And ԤԼ˳��� Is Null And Nvl(�Ƿ�ԤԼ,0)=1 Order By ���,��ʼʱ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Sum(Nvl(����,0)) As ������λ����,���  From �ٴ�����Һſ��Ƽ�¼ Where ��¼id=[1] And ����=1 And ���Ʒ�ʽ=3 Group By ���"
        Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Count(1) As �ѹ�����,��� From �ٴ�������ſ��� Where ��¼ID=[1] And ԤԼ˳��� Is Null And Nvl(�Һ�״̬,0) <> 0 Group By ���"
        Set rsUsed = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!����)) <> 0 Then
                rsUnit.Filter = "���=" & Val(rsTemp!���)
                If rsUnit.EOF Then
                    lng������λ���� = 0
                Else
                    lng������λ���� = Val(Nvl(rsUnit!������λ����))
                End If
                rsUsed.Filter = "���=" & Val(rsTemp!���)
                If rsUsed.EOF Then
                    lng�ѹ����� = 0
                Else
                    lng�ѹ����� = Val(Nvl(rsUsed!�ѹ�����))
                End If
                With vsfList
                    If intCurrentTime = -1 Then
                        intCurrentTime = Val(Split(Nvl(rsTemp!��ʼʱ��), ":")(0))
                        .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                        intCol = intCol + 1
                    Else
                        If intCurrentTime = Val(Split(Nvl(rsTemp!��ʼʱ��), ":")(0)) Then
                            intCol = intCol + 1
                        Else
                            .Rows = .Rows + 1
                            intCurrentTime = Val(Split(Nvl(rsTemp!��ʼʱ��), ":")(0))
                            .TextMatrix(.Rows - 1, 0) = Format(intCurrentTime, "00") & ":00"
                            intCol = 1
                        End If
                    End If
                    
                    If intCol >= .Cols Then .Cols = .Cols + 1
                    If vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) <> "" And _
                       Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss") >= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")), "yyyy-mm-dd hh:mm:ss") And _
                       Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss") <= Format(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("������ֹʱ��")), "yyyy-mm-dd hh:mm:ss") Then
                        .TextMatrix(.Rows - 1, intCol) = Format(Nvl(rsTemp!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "hh:mm") & vbCrLf & "ԤԼ" & Val(Nvl(rsTemp!����)) - lng������λ���� - lng�ѹ����� & "��(��)"
                    Else
                        .TextMatrix(.Rows - 1, intCol) = Format(Nvl(rsTemp!��ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsTemp!��ֹʱ��), "hh:mm") & vbCrLf & "ԤԼ" & Val(Nvl(rsTemp!����)) - lng������λ���� - lng�ѹ����� & "��"
                    End If
                    
                    .Cell(flexcpData, .Rows - 1, intCol) = Nvl(rsTemp!���)
                    If Val(Nvl(rsTemp!����)) - lng������λ���� - lng�ѹ����� = 0 Then
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                    If Val(Nvl(rsTemp!�Ƿ�ԤԼ)) = 0 Then
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                    If CDate(Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlngԤԼ��Чʱ��, datNow) Then
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                    If Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                    End If
                End With
            End If
            rsTemp.MoveNext
        Loop
    End If
    With vsfList
        For i = 0 To .Rows - 1
            .RowHeight(i) = 500
            .Cell(flexcpFontBold, i, 0) = True
            .Cell(flexcpFontSize, i, 0) = 20
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 1500
            If i = 0 Then
                .ColAlignment(i) = flexAlignCenterTop
            Else
                .ColAlignment(i) = flexAlignCenterCenter
            End If
        Next i
    End With
    blnFind = False
    With vsfList
        For i = 0 To .Rows - 1
            If blnFind = False Then
                For j = 1 To .Cols - 1
                    If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False And vsfList.TextMatrix(i, j) <> "" Then
                        .Select i, j
                        Call vsfList_EnterCell
                        blnFind = True
                        Exit For
                    End If
                Next j
            End If
        Next i
        mblnNotClick = True
        If blnFind = False Then .Select 0, 0
        mblnNotClick = False
    End With
    vsfList.Redraw = flexRDDirect
End Sub

Private Sub vsfPlan_GotFocus()
    With vsfPlan
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = 16772055
    End With
End Sub

