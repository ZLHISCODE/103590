VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.8#0"; "ZLIDKIND.OCX"
Begin VB.Form frmServiceChangeNum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ԤԼ����"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10950
   Icon            =   "frmServiceChangeNum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   300
      Left            =   705
      TabIndex        =   38
      Top             =   495
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "����"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1380
      Left            =   15
      ScaleHeight     =   1320
      ScaleWidth      =   10830
      TabIndex        =   0
      Top             =   1365
      Width           =   10890
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   -45
         TabIndex        =   39
         Top             =   1290
         Width           =   11145
      End
      Begin VB.TextBox txtNO 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   540
         Width           =   1530
      End
      Begin VB.TextBox txtDoc 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   1530
      End
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3225
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   945
         Width           =   1530
      End
      Begin VB.TextBox txtAdd 
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   945
         Width           =   3240
      End
      Begin VB.TextBox txtFee 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   945
         Width           =   1110
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   -75
         TabIndex        =   25
         Top             =   480
         Width           =   11115
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   45
         Width           =   1530
      End
      Begin VB.TextBox txtGender 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   45
         Width           =   570
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   45
         Width           =   810
      End
      Begin VB.TextBox txtPhone 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   45
         Width           =   1110
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   45
         Width           =   3240
      End
      Begin VB.TextBox txtAppTime 
         Enabled         =   0   'False
         Height          =   330
         Left            =   3225
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   1530
      End
      Begin VB.TextBox txtRegNO 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   330
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   3240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   2835
         TabIndex        =   35
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         Height          =   180
         Left            =   6765
         TabIndex        =   33
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�Һŷ�"
         Height          =   180
         Left            =   5025
         TabIndex        =   28
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   375
         TabIndex        =   26
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   390
         TabIndex        =   24
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2475
         TabIndex        =   23
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   3570
         TabIndex        =   22
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��ϵ�绰"
         Height          =   180
         Left            =   4845
         TabIndex        =   21
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "��סַ"
         Height          =   180
         Left            =   6945
         TabIndex        =   20
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼʱ��"
         Height          =   180
         Left            =   2475
         TabIndex        =   19
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����"
         Height          =   180
         Left            =   30
         TabIndex        =   18
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   5205
         TabIndex        =   17
         Top             =   615
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   7125
         TabIndex        =   16
         Top             =   615
         Width           =   360
      End
   End
   Begin VB.PictureBox picReg 
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   120
      ScaleHeight     =   6180
      ScaleWidth      =   9015
      TabIndex        =   29
      Top             =   2865
      Width           =   9015
      Begin VB.PictureBox picSplit 
         BorderStyle     =   0  'None
         Height          =   100
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   3855
         TabIndex        =   41
         Top             =   2565
         Width           =   3855
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   2730
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   75
         Width           =   1125
      End
      Begin VB.PictureBox picTime 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4005
         ScaleHeight     =   330
         ScaleWidth      =   2745
         TabIndex        =   31
         Top             =   60
         Visible         =   0   'False
         Width           =   2745
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   315
            Left            =   750
            TabIndex        =   4
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   155320322
            CurrentDate     =   42340
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "ԤԼʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   32
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Left            =   465
         TabIndex        =   2
         Top             =   68
         Width           =   1320
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
         Height          =   2100
         Left            =   60
         TabIndex        =   6
         Top             =   465
         Width           =   3360
         _cx             =   5927
         _cy             =   3704
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
         Rows            =   2
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceChangeNum.frx":058A
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   3120
         Left            =   60
         TabIndex        =   7
         Top             =   2670
         Width           =   5925
         _cx             =   10451
         _cy             =   5503
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
         HighLight       =   1
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
         FormatString    =   $"frmServiceChangeNum.frx":08BD
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
      Begin XtremeCalendarControl.DatePicker dtpMain 
         Height          =   2100
         Left            =   5190
         TabIndex        =   5
         Top             =   450
         Width           =   3060
         _Version        =   1048579
         _ExtentX        =   5397
         _ExtentY        =   3704
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowTodayButton =   0   'False
         ShowNoneButton  =   0   'False
         Show3DBorder    =   0
         MaxSelectionCount=   1
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼʱ��"
         Height          =   180
         Left            =   1980
         TabIndex        =   40
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   60
         TabIndex        =   30
         Top             =   135
         Width           =   360
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   7545
      Top             =   3675
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmServiceChangeNum.frx":09C9
      Left            =   2445
      Top             =   2415
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceChangeNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrʱ��s As String
Private mbln�Ƿ��շ� As Boolean
Private mdbl�շѽ�� As Double
Private mblnUnload As Boolean, mblnNotClick As Boolean
Public mlng��ϢID As Long
Private mstrName As String, mstrGender As String
Private mstrAge As String, mlngԤԼ��Чʱ�� As Long
Private mstrPhone As String
Private mstrNo As String
Private mstrAppTime As String
Private mstrInfo As String
Private mstrRegNo As String
Private mstrDoc As String, mblnInit As Boolean, mblnAppointmentChange As Boolean
Private mblnChangeByCode As Boolean, mlngRow As Long, msngTime As Single
Private mstrPriceGrade As String

Public Sub ShowMe(frmMain As Object)
    Me.Show 1, frmMain
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
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
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    cbsThis.ActiveMenuBar.Visible = False
    
    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, 3839, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 145, 80, DockTopOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Handle = picInfo.Hwnd
    objPane.MaxTrackSize.Height = 88
    objPane.MinTrackSize.Height = 88
    
    Set objPane = dkpMain.CreatePane(3, 145, 120, DockBottomOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Title = "���ﰲ����Ϣ"
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

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 3839 '����
            If SaveData = True Then Unload Me
        Case Else
            Unload Me
    End Select
End Sub

Private Function SaveData() As Boolean
    '�������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim datApp As Date
    Dim strTemp As String
    
    If vsfPlan.RowData(vsfPlan.Row) = "" Or vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����")) = "" Then
        MsgBox "��ѡ��һ�����Ž��л���!", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsfPlan
        strTemp = .TextMatrix(.Row, .ColIndex("����")) & "-" & .TextMatrix(.Row, .ColIndex("����")) & _
                IIf(.TextMatrix(.Row, .ColIndex("ҽ��")) = "", "", "(" & .TextMatrix(.Row, .ColIndex("ҽ��")) & ")")
    End With
    
    If MsgBox("�Ƿ�ȷ��������:" & txtRegNO.Text & "-" & txtDept.Text & _
                IIf(txtDoc.Text = "", "", "(" & txtDoc.Text & ")") & " ���ﵽ" & vbCrLf & _
                "����:" & strTemp & "?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
        Exit Function
    End If
    
    If CheckLimit(Val(vsfPlan.RowData(vsfPlan.Row))) = False Then Exit Function
    
    If vsfList.Visible Then
        If vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col) = "" Then
            MsgBox "��ѡ��һ����Ч����Ž��л���!", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʱ��"))) = 1 Then
            If Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) <> 0 Then
                strSQL = "Select ��ʼʱ�� From �ٴ�������ſ��� Where ��¼ID=[1] And ���=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)))
                If Not rsTemp.EOF Then
                    datApp = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd hh:mm:ss"))
                Else
                    datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                End If
            Else
                datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
            End If
        Else
            datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
        End If
    Else
        datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
    End If
    
    
    If datApp < DateAdd("n", -1 * mlngԤԼ��Чʱ��, zlDatabase.Currentdate) Then
        MsgBox "ԤԼʱ��С���˿�ԤԼʱ��(" & Format(DateAdd("n", -1 * mlngԤԼ��Чʱ��, zlDatabase.Currentdate), "hh:mm:ss") & "),�޷�ԤԼ!", vbInformation, gstrSysName
        Exit Function
    End If
    If Not (Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʱ��"))) = 1 And Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ſ���"))) = 1) Then
        If Check��Чʱ���(Val(vsfPlan.RowData(vsfPlan.Row)), datApp) = False Then
            MsgBox "��ǰѡ��ĳ����¼��" & Format(datApp, "yyyy-mm-dd hh:mm:ss") & "������,������Һ�ʱ��!", vbInformation, gstrSysName
            If dtpTime.Enabled And dtpTime.Visible Then dtpTime.SetFocus
            Exit Function
        End If
    End If
    
    strSQL = "Select Zl_�ٴ���������_Check([1],[2],[3]) As �����Լ�� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), txtGender.Text, txtAge.Text)
    If rsTemp.EOF Then
        MsgBox "��ǰѡ��Ĳ��˲����øúű�!", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Mid(Nvl(rsTemp!�����Լ��), 1, 1)) <> 0 Then
            MsgBox "��ǰѡ��Ĳ��˲����øúű�!" & vbCrLf & "ԭ��:" & Mid(Nvl(rsTemp!�����Լ��), InStr(Nvl(rsTemp!�����Լ��), "|") + 1), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select 1 From �ٴ������¼ Where Id=[1] And [2] Between ��ʼʱ�� And ��ֹʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsfPlan.RowData(vsfPlan.Row), datApp)
    If rsTemp.EOF Then
        MsgBox "��ǰԤԼʱ�䲻�ڸú������Чʱ����,������ѡ��ԤԼʱ��!", vbInformation, gstrSysName
        Exit Function
    End If
    strSQL = "Zl_���߷�������_����(" & mlng��ϢID & ",'"
    strSQL = strSQL & Trim(txtNO.Text) & "',"
    If vsfList.Visible Then
        strSQL = strSQL & Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) & ","
    Else
        strSQL = strSQL & "Null,"
    End If
    strSQL = strSQL & "To_Date('" & datApp & "','yyyy-mm-dd hh24:mi:ss')" & ","
    strSQL = strSQL & vsfPlan.RowData(vsfPlan.Row) & ","
    strSQL = strSQL & "'" & txtAdd.Text & "',"
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    strSQL = strSQL & "'" & mstrPriceGrade & "')" '�۸�ȼ�
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
End Function

Private Sub dtpTime_Change()
    Dim str���� As String, i As Integer, lngRow As Long
    Dim str����ʱ�� As String
    If Not dtpMain.Visible Then Exit Sub
    If Not dtpMain.Enabled Then Exit Sub
    
    str���� = Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-MM-dd")

    If str���� = "" Then str���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    str����ʱ�� = str���� & " " & Format(dtpTime.Value, "hh:mm:00")
    lngRow = 0
    If CDate(str����ʱ��) > CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ֹʱ��"))) Then
        '����ʱ��İ��ţ�����Ѱ�Ҷ�λ
        For i = 1 To vsfPlan.Rows - 1
            With vsfPlan
                If .TextMatrix(.Row, .ColIndex("����")) = .TextMatrix(i, .ColIndex("����")) And _
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
                If .TextMatrix(.Row, .ColIndex("����")) = .TextMatrix(i, .ColIndex("����")) And _
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

Private Function Check��Чʱ���(lng��¼ID As Long, datTime As Date) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    With vsfPlan
        '��ſ��Ʒ�ʱ�κ�,��������ʱ���Ƿ��ڳ����¼ʱ����
        If Val(.TextMatrix(.Row, .ColIndex("��ſ���"))) = 1 And Val(.TextMatrix(.Row, .ColIndex("��ʱ��"))) = 1 Then
            strSQL = "Select 1 From �ٴ������¼ Where ID=[1] And [2] Between Nvl(ͣ�￪ʼʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(ͣ����ֹʱ��,To_Date('3000-01-01', 'yyyy-mm-dd')) "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��¼ID, datTime)
            If rsTemp.EOF Then
                Check��Чʱ��� = True
            Else
                Check��Чʱ��� = False
            End If
            Exit Function
        End If
    End With
    
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

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 3839
        Control.Enabled = vsfPlan.RowData(vsfPlan.Row) <> ""
    End Select
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub dtpMain_SelectionChanged()
    Dim datNow As Date
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd") > Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
        MsgBox "���ܻ��ﵱǰʱ��֮ǰ�İ���!", vbInformation, gstrSysName
        dtpMain.SelectRange datNow, datNow
        dtpMain.Select datNow
        dtpMain.EnsureVisibleSelection
        dtpMain.RedrawControl
    End If
    Call LoadPlan
End Sub

Private Sub InitGrid()
    Dim i As Integer
    With vsfPlan
        .MergeRow(0) = True
        .Rows = 2
        For i = 0 To .Rows - 1
            .RowHeight(i) = 350
        Next i
    End With
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call dtpTime_Change
End Sub

Private Sub Form_Activate()
    mblnInit = True
    If mblnUnload Then
        mblnUnload = False
        Unload Me
        Exit Sub
    End If
    Call vsfPlan_EnterCell
    If txtFilter.Visible And txtFilter.Enabled Then txtFilter.SetFocus
    mblnInit = False
End Sub
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picInfo.Hwnd
    Case 3
        Item.Handle = picReg.Hwnd
    End Select
End Sub
Private Sub InitPara()
'    mlngԤԼ��Чʱ�� = Val(Split(zlDatabase.GetPara("ԤԼ����ʱ��", glngSys, 1111, "1|60") & "|", "|")(1))
    mlngԤԼ��Чʱ�� = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer, strSQL As String, rsTemp As ADODB.Recordset
    Dim datTime As Date
    mblnInit = True
    If mstrInfo = "" Then
        SetWithValue
    Else
        Call SetValue
    End If
    Set txtAdd.Container = Me
    txtAdd.Left = txtAdd.Left + 15
    txtAdd.Top = txtAdd.Top + 540
    txtAdd.Locked = False
    Call DefMainCommandBars
    Call InitPanel
    Call InitGrid
    Call InitPara
    dtpTime.Value = Now
    If Format(txtAppTime.Text, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
        datTime = CDate(zlDatabase.Currentdate)
    Else
        dtpTime = CDate(txtAppTime.Text)
    End If
    
    dtpMain.SelectRange dtpTime, dtpTime
    dtpMain.Select dtpTime
    dtpMain.EnsureVisibleSelection
    dtpMain.EnsureVisible dtpTime
    dtpMain.RedrawControl
    
    dtpMain.HighlightToday = False
    dtpMain.ShowNonMonthDays = False
    dtpMain.PaintManager.ControlBackColor = &H8000000F
    dtpMain.PaintManager.DayBackColor = &H8000000F
    dtpMain.PaintManager.DaysOfWeekBackColor = &H8000000F
    
    strSQL = "Select 1 From ������ü�¼ Where NO=[1] And ��¼����=4 And ����ID Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtNO.Text))
    If Not rsTemp.EOF Then
        mbln�Ƿ��շ� = False
    Else
        mbln�Ƿ��շ� = True
        strSQL = "Select Sum(Nvl(���ʽ��, 0)) As ���" & vbNewLine & _
                "From ������ü�¼" & vbNewLine & _
                "Where NO = [1] And ��¼���� = 4 And" & vbNewLine & _
                "      �շ�ϸĿid Not In" & vbNewLine & _
                "      (Select �շ�ϸĿid" & vbNewLine & _
                "       From �շ��ض���Ŀ" & vbNewLine & _
                "       Where �ض���Ŀ = '������'" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select ����id From �շѴ�����Ŀ Where ����id In (Select �շ�ϸĿid From �շ��ض���Ŀ Where �ض���Ŀ = '������'))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtNO.Text))
        mdbl�շѽ�� = Val(Nvl(rsTemp!���))
    End If
    mblnUnload = False
    Call LoadPlan
    IDKind.RaisEffect picInfo, -1
    mblnInit = False
End Sub

Public Sub InitValue(strName As String, strGender As String, strAge As String, strPhone As String, _
    strNO As String, strAppTime As String, strInfo As String, ByVal strPriceGrade As String)
    mstrName = strName
    mstrGender = strGender
    mstrAge = strAge
    mstrPhone = strPhone
    mstrNo = strNO
    mstrAppTime = strAppTime
    mstrInfo = strInfo
    mstrPriceGrade = strPriceGrade '�۸�ȼ�
End Sub

Public Sub InitWithValue(strName As String, strGender As String, strAge As String, strPhone As String, _
    strNO As String, strAppTime As String, strRegNo As String, strDoc As String, ByVal strPriceGrade As String)
    mstrName = strName
    mstrGender = strGender
    mstrAge = strAge
    mstrPhone = strPhone
    mstrNo = strNO
    mstrAppTime = strAppTime
    mstrRegNo = strRegNo
    mstrDoc = strDoc
    mstrPriceGrade = strPriceGrade '�۸�ȼ�
End Sub

Private Sub SetValue()
    Dim strArray() As String
    With Me
        .txtName = mstrName
        .txtGender = mstrGender
        .txtAge = mstrAge
        .txtPhone = mstrPhone
        .txtNO = mstrNo
        .txtAppTime = mstrAppTime
        strArray = Split(mstrInfo, "   ")
        .txtRegNO = Split(strArray(0), ":")(1) '& "[" & Split(strArray(1), ":")(1) & "]"
        .txtItem = Split(strArray(3), ":")(1)
        .txtDept = Split(strArray(2), ":")(1)
        .txtDoc = Split(strArray(4), ":")(1)
    End With
End Sub

Private Sub SetWithValue()
    Dim strArray() As String
    With Me
        .txtName = mstrName
        .txtGender = mstrGender
        .txtAge = mstrAge
        .txtPhone = mstrPhone
        .txtNO = mstrNo
        .txtAppTime = mstrAppTime
        .txtRegNO = mstrRegNo
        .txtDoc = mstrDoc
    End With
End Sub

Private Sub LoadPlan()
    Dim strSQL As String, rsPlan As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim datApp As Date, i As Integer, dblMoney As Double, blnAdd As Boolean
    Dim strTime() As String, lngLeft As Long
    Dim dbl��Ŀ��� As Double
    
    datApp = dtpMain.Selection.Blocks(0).DateBegin
    strSQL = "Select a.id, b.����, b.���� As ����, c.���� As ����, c.���� As ���Ҽ���, b.����Id, a.�ϰ�ʱ�� As ʱ��, " & _
            "        d.���� As ��Ŀ, zlSpellcode(d.����) As ��Ŀ����, a.����ҽ������, a.����ҽ��ID, a.ҽ��Id, a.ҽ������ As ҽ��, e.���� As ҽ������, " & _
            "        a.��Ŀid, a.��Լ�� As ��Լ, a.��Լ�� As ��Լ, Nvl(a.�Ƿ��ʱ��,0) As ��ʱ��, Nvl(a.�Ƿ���ſ���,0) As ��ſ���, " & _
            "        a.��������, a.ȱʡԤԼʱ�� , a.���￪ʼʱ�� , a.������ֹʱ�� , a.��ʼʱ��, a.��ֹʱ�� " & vbNewLine & _
            "From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D, ��Ա�� E" & vbNewLine & _
            "Where a.��Դid = b.Id  And Nvl(C.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.��Ŀid = d.Id And b.����id = c.Id And (c.վ�� Is Null Or c.վ�� = '" & gstrNodeNo & "') And (a.�������� = [1] Or a.�������� = [2]) And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��,a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��,a.��ʼʱ��) Or Exists (Select 1 From �ٴ�������ſ��� C,�ٴ������¼ D Where D.ID=A.ID And C.��¼ID=D.ID And Nvl(C.�Ƿ�ͣ��,0) = 0 And D.�Ƿ���ſ��� =1 And D.�Ƿ��ʱ�� = 1 And C.��ʼʱ�� <> C.��ֹʱ��)) " & _
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
    vsfPlan.Clear 1
    vsfPlan.Rows = 2
    Do While Not rsPlan.EOF
        blnAdd = True
        
        dbl��Ŀ��� = Get��Ŀ���(Val(Nvl(rsPlan!��ĿID)), mstrPriceGrade)
        If mbln�Ƿ��շ� Then
            '�շѵĹҺż�¼,��������ͬ�Ĳ��ܻ���
            If RoundEx(mdbl�շѽ��, 6) <> RoundEx(dbl��Ŀ���, 6) Then
                blnAdd = False
            End If
        End If
        
        If blnAdd Then
            With vsfPlan
                .RowData(.Rows - 1) = Val(Nvl(rsPlan!ID))
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsPlan!����)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsPlan!����)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsPlan!����)
                .TextMatrix(.Rows - 1, .ColIndex("ʱ��")) = Nvl(rsPlan!ʱ��)
                If InStr("," & mstrʱ��s & ",", "," & Nvl(rsPlan!ʱ��) & ",") = 0 Then
                    mstrʱ��s = mstrʱ��s & "," & Nvl(rsPlan!ʱ��)
                End If
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = Nvl(rsPlan!��Ŀ)
                .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(dbl��Ŀ���, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("��Լ")) = Nvl(rsPlan!��Լ)
                .TextMatrix(.Rows - 1, .ColIndex("��Լ")) = Nvl(rsPlan!��Լ)
                .TextMatrix(.Rows - 1, .ColIndex("��ʱ��")) = Nvl(rsPlan!��ʱ��)
                .TextMatrix(.Rows - 1, .ColIndex("��ſ���")) = Nvl(rsPlan!��ſ���)
                .TextMatrix(.Rows - 1, .ColIndex("ҽ��")) = Nvl(rsPlan!ҽ��)
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(Nvl(rsPlan!��������), "yyyy-mm-dd")
                .TextMatrix(.Rows - 1, .ColIndex("ҽ������")) = Nvl(rsPlan!ҽ������)
                .TextMatrix(.Rows - 1, .ColIndex("���Ҽ���")) = Nvl(rsPlan!���Ҽ���)
                .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = Nvl(rsPlan!��Ŀ����)
                .TextMatrix(.Rows - 1, .ColIndex("ԤԼʱ��")) = Format(Nvl(rsPlan!ȱʡԤԼʱ��), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = Format(Nvl(rsPlan!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = Format(Nvl(rsPlan!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
                If Nvl(rsPlan!����ҽ������) <> "" Then
                    .Cell(flexcpData, .Rows - 1, .ColIndex("����ҽ��")) = Nvl(rsPlan!����ҽ������) & "(" & Format(Nvl(rsPlan!���￪ʼʱ��), "hh:mm") & "-" & Format(Nvl(rsPlan!������ֹʱ��), "hh:mm") & ")"
                    .TextMatrix(.Rows - 1, .ColIndex("����ҽ��")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("����ҽ������")) = Nvl(rsPlan!����ҽ������)
                    .TextMatrix(.Rows - 1, .ColIndex("����ҽ��ID")) = Nvl(rsPlan!����ҽ��id)
                    .TextMatrix(.Rows - 1, .ColIndex("���￪ʼʱ��")) = Nvl(rsPlan!���￪ʼʱ��)
                    .TextMatrix(.Rows - 1, .ColIndex("������ֹʱ��")) = Nvl(rsPlan!������ֹʱ��)
                End If
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
        vsfPlan.Redraw = flexRDDirect
        vsfList.Visible = False
        picSplit.Visible = False
        vsfPlan.Height = 4915
        vsfPlan.Select 1, 1
        Exit Sub
    End If
    Call ShowRow
    If vsfPlan.Rows <> 2 Then vsfPlan.Rows = vsfPlan.Rows - 1
    For i = 0 To vsfPlan.Rows - 1
        vsfPlan.RowHeight(i) = 322
    Next i
'    vsfPlan.AutoSize 0, vsfPlan.Cols - 1
    zl_vsGrid_Para_Restore 1115, vsfPlan, Me.Name, "vsfPlan"
    vsfPlan.Redraw = flexRDDirect
    If vsfPlan.TextMatrix(1, vsfPlan.ColIndex("����")) = "" Then
'        MsgBox "��ǰԤԼ��¼û�п��Ի���ļ�¼,�޷�����!", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save 1115, vsfPlan, Me.Name, "vsfPlan"
End Sub

Private Sub picReg_Resize()
    On Error Resume Next
    With dtpMain
        .Left = picReg.ScaleWidth - .Width - 15
    End With
    With vsfPlan
        .Width = dtpMain.Left - 60
    End With
    With vsfList
        .Width = picReg.ScaleWidth - 150
        .Height = picReg.ScaleHeight - 2600
    End With
    picSplit.Width = picReg.ScaleWidth
End Sub

Private Sub ShowRow()
    Dim i As Integer, blnHide As Boolean
    Dim blnEnable As Boolean, strTimeRange As String
    If cboTime.Text <> "����" Then strTimeRange = cboTime.Text
    With vsfPlan
        For i = 1 To .Rows - 1
            blnHide = False
            If txtFilter <> "" Then
                blnHide = True
                If .TextMatrix(i, .ColIndex("����")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("����")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("����")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("��Ŀ")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If .TextMatrix(i, .ColIndex("ҽ��")) Like "*" & txtFilter.Text & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("���Ҽ���"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("ҽ������"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
                If UCase(.TextMatrix(i, .ColIndex("��Ŀ����"))) Like "*" & UCase(txtFilter.Text) & "*" Then blnHide = False
            End If
            If cboTime.Text <> "����" Then
                If InStr(strTimeRange & ",", .TextMatrix(i, .ColIndex("ʱ��"))) = 0 Then blnHide = True
            End If
            .RowHidden(i) = blnHide
        Next i
    End With
    blnEnable = False
    With vsfPlan
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                .Select i, 1
                blnEnable = True
                Call vsfPlan_EnterCell
                Exit For
            End If
        Next i
        If blnEnable = False Then
            picTime.Visible = True
            vsfList.Visible = False
            picSplit.Visible = False
            vsfPlan.Height = 4915
        End If
    End With
End Sub

Private Sub txtFilter_Change()
    Call ShowRow
End Sub

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

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnNotClick = False Then
        With vsfList
            If .TextMatrix(NewRow, NewCol) = "" Then Cancel = True
            If .Cell(flexcpFontStrikethru, NewRow, NewCol) = True Then Cancel = True
            If .Cell(flexcpForeColor, NewRow, NewCol) <> vbBlack And .Cell(flexcpForeColor, NewRow, NewCol) <> 2 Then Cancel = True
        End With
    End If
End Sub

Private Sub vsfList_DblClick()
'    vsfList.Select vsfList.Row, vsfList.Col
    If SaveData = True Then Unload Me
End Sub

Private Sub vsfList_EnterCell()
    If vsfList.Row >= vsfList.Rows Then Exit Sub
    If vsfList.Col >= vsfList.Cols Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), ":") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "-") = 0 Then Exit Sub
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "ԤԼ") > 0 Then
        dtpTime.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(0), "-")(0)
    Else
        dtpTime.Value = Split(Split(vsfList.TextMatrix(vsfList.Row, vsfList.Col), vbCrLf)(1), "-")(0)
    End If
    If InStr(vsfList.TextMatrix(vsfList.Row, vsfList.Col), "��") > 0 Then
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
    Else
        vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
    End If
    
End Sub

Private Sub vsfplan_DblClick()
    If SaveData = True Then Unload Me
End Sub

Private Sub vsfPlan_EnterCell()
    Dim i As Integer, j As Integer
    Dim sngTime As Single, datApp As Date
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnFind As Boolean
    If mblnInit = False Then
        If mblnChangeByCode Then Exit Sub
        sngTime = Timer
        If Format(sngTime, "0.000") - Format(msngTime, "0.000") < 0.1 Then
            mblnChangeByCode = True
            If mlngRow <> 0 Then vsfPlan.Select mlngRow, 0
            mblnChangeByCode = False
            Exit Sub
        End If
        msngTime = Timer
        mlngRow = vsfPlan.Row
    End If
    With vsfPlan
        If Val(.TextMatrix(.Row, .ColIndex("��ʱ��"))) = 1 Then
            With vsfList
                For i = 0 To .Rows - 1
                    .RowHeight(i) = 500
                    .Cell(flexcpFontBold, i, 0) = True
                    .Cell(flexcpFontSize, i, 0) = 16
                Next i
            End With
            picTime.Visible = False
            picSplit.Visible = True
            vsfList.Visible = True
            vsfPlan.Height = picSplit.Top - vsfPlan.Top - 15
            Call LoadTimePlan
            blnFind = False
            With vsfList
                For i = 0 To .Rows - 1
                    If blnFind = False Then
                        For j = 1 To .Cols - 1
                            If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False Then
                                .Select i, j
                                blnFind = True
                                Exit For
                            End If
                        Next j
                    End If
                Next i
            End With
        Else
            If mblnAppointmentChange = False Then
                If vsfPlan.TextMatrix(.Row, .ColIndex("ԤԼʱ��")) <> "" And IsDate(vsfPlan.TextMatrix(.Row, .ColIndex("ԤԼʱ��"))) = True Then
                    If Format(vsfPlan.TextMatrix(.Row, .ColIndex("��������")), "yyyy-mm-dd") <> Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") Then
                        dtpTime.Value = vsfPlan.TextMatrix(.Row, .ColIndex("��ֹʱ��"))
                    Else
                        dtpTime.Value = vsfPlan.TextMatrix(.Row, .ColIndex("ԤԼʱ��"))
                    End If
                Else
                    dtpTime.Value = zlDatabase.Currentdate
                End If
            End If
            If Val(vsfPlan.TextMatrix(.Row, .ColIndex("��ſ���"))) = 0 Then
                picTime.Visible = True
                vsfList.Visible = False
                picSplit.Visible = False
                vsfPlan.Height = 4915
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
                picTime.Visible = True
                vsfList.Visible = True
                picSplit.Visible = True
                vsfPlan.Height = picSplit.Top - vsfPlan.Top - 15
                Call LoadSerialPlan
                blnFind = False
                With vsfList
                    For i = 0 To .Rows - 1
                        If blnFind = False Then
                            For j = 1 To .Cols - 1
                                If (vsfList.Cell(flexcpForeColor, i, j) = vbBlack Or vsfList.Cell(flexcpForeColor, i, j) = 2) And vsfList.Cell(flexcpFontStrikethru, i, j) = False Then
                                    .Select i, j
                                    blnFind = True
                                    Exit For
                                End If
                            Next j
                        End If
                    Next i
                End With
            End If
            If vsfList.Visible Then
                If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ʱ��"))) = 1 Then
                    If Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)) <> 0 Then
                        strSQL = "Select ��ʼʱ�� From �ٴ�������ſ��� Where ��¼ID=[1] And ���=[2]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)), Val(vsfList.Cell(flexcpData, vsfList.Row, vsfList.Col)))
                        If Not rsTemp.EOF Then
                            datApp = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd hh:mm:ss"))
                        Else
                            datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                        End If
                    Else
                        datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                    End If
                Else
                    datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
                End If
            Else
                datApp = CDate(Format(dtpMain.Selection.Blocks(0).DateBegin, "yyyy-mm-dd") & " " & Format(dtpTime.Value, "hh:mm:00"))
            End If
            If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��")) <> "" Then
                If datApp >= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��"))) And datApp <= CDate(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("���￪ʼʱ��"))) Then
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = ""
                Else
                    vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("����ҽ��")) = vsfPlan.Cell(flexcpData, vsfPlan.Row, vsfPlan.ColIndex("����ҽ��"))
                End If
            End If
        End If
    End With
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
    
    strSQL = "Select ���, ��ʼʱ��, ��ֹʱ��, �Ƿ�ԤԼ, �Һ�״̬, ����, ���� From �ٴ�������ſ��� Where ��¼id = [1] Order By ���,��ʼʱ��"
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
                .Cell(flexcpFontSize, i, j) = 12
            Next j
        Next i
        For i = 0 To .Cols - 1
            .ColWidth(i) = 1000
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
    vsfList.RowHidden(0) = True
    vsfList.Redraw = flexRDDirect
End Sub

Private Sub LoadTimePlan()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intCurrentTime As Integer, intCol As Integer, datNow As Date
    Dim rsUnit As ADODB.Recordset, lng������λ���� As Long, lng�ѹ����� As Long
    Dim i As Long, datTime As Date, blnFind As Boolean, j As Integer
    Dim rsTmp As ADODB.Recordset
    vsfList.Redraw = flexRDNone
    vsfList.Clear
    vsfList.Rows = 1
    vsfList.Cols = 2
    vsfList.FixedRows = 0
    vsfList.FixedCols = 1
    intCol = 0
    datTime = dtpMain.Selection.Blocks(0).DateBegin
    datNow = zlDatabase.Currentdate
    intCurrentTime = -1
    If Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("��ſ���"))) = 1 Then
        strSQL = "Select ���, To_Char(��ʼʱ��,'hh24:mi:ss') As ��ʼʱ��, ��ʼʱ�� As ���ʱ��, To_Char(��ֹʱ��,'hh24:mi:ss') As ��ֹʱ��, �Ƿ�ԤԼ, �Һ�״̬, ����, ���� From �ٴ�������ſ��� Where ��¼id = [1] And ��ʼʱ�� <> ��ֹʱ�� Order By ���,��ʼʱ��"
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
                If CDate(Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlngԤԼ��Чʱ��, datNow) Then
                    .Cell(flexcpForeColor, .Rows - 1, intCol) = &H8000000C
                End If
            End With
            rsTemp.MoveNext
        Loop
    Else
        strSQL = "Select ���, To_Char(��ʼʱ��,'hh24:mi:ss') As ��ʼʱ��, ��ʼʱ�� As ���ʱ��, To_Char(��ֹʱ��,'hh24:mi:ss') As ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ���� From �ٴ�������ſ��� Where ��¼id = [1] And Nvl(�Ƿ�ԤԼ,0) = 1 And ԤԼ˳��� Is Null Order By ���,��ʼʱ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Sum(Nvl(����,0)) As ������λ����,��� From �ٴ�����Һſ��Ƽ�¼ Where ��¼id=[1] And ����=1 And ���Ʒ�ʽ=3 Group By ���"
        Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        strSQL = "Select Count(1) As �ѹ�����,��� From �ٴ�������ſ��� Where ��¼ID=[1] And ԤԼ˳��� Is Null And Nvl(�Һ�״̬,0) <> 0 Group By ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsfPlan.RowData(vsfPlan.Row)))
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!����)) <> 0 Then
                rsUnit.Filter = "���=" & Val(rsTemp!���)
                If rsUnit.EOF Then
                    lng������λ���� = 0
                Else
                    lng������λ���� = Val(Nvl(rsUnit!������λ����))
                End If
                rsTmp.Filter = "���=" & Val(rsTemp!���)
                If rsTmp.EOF Then
                    lng�ѹ����� = 0
                Else
                    lng�ѹ����� = Val(Nvl(rsTmp!�ѹ�����))
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
'                    Select Case Val(Nvl(rsTemp!�Һ�״̬))
'                        Case 0
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
'                        Case 1 '�ѹ�
'                            .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
'                        Case 2
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbGreen
'                        Case 3
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlue
'                        Case 4
'                            .Cell(flexcpForeColor, .Rows - 1, intCol) = vbRed
'                        Case 5
'                            .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
'                    End Select
                    If Val(Nvl(rsTemp!����)) - lng������λ���� - lng�ѹ����� = 0 Then
                        .Cell(flexcpFontStrikethru, .Rows - 1, intCol) = True
                        .Cell(flexcpForeColor, .Rows - 1, intCol) = vbBlack
                    End If
                    If CDate(Format(Nvl(rsTemp!���ʱ��), "yyyy-mm-dd hh:mm:ss")) < DateAdd("n", -1 * mlngԤԼ��Чʱ��, datNow) Then
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

Private Sub cboTime_Click()
    If mblnNotClick Then Exit Sub
    Call ShowRow
End Sub

Private Sub cboTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub vsfPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPlan
        If OldRow < .Rows Then
            If OldRow Mod 2 = 1 Then
                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEEEEEE
            End If
        End If
        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = 16772055
    End With
End Sub

