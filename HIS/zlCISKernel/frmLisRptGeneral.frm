VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLisRptGeneral 
   BorderStyle     =   0  'None
   Caption         =   "frmLisStationWrite"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12540
   Icon            =   "frmLisRptGeneral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTab 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   8775
      ScaleHeight     =   2280
      ScaleWidth      =   3900
      TabIndex        =   33
      Top             =   3195
      Width           =   3900
      Begin XtremeSuiteControls.TabControl TabThis 
         Height          =   2280
         Left            =   75
         TabIndex        =   34
         Top             =   165
         Width           =   3765
         _Version        =   589884
         _ExtentX        =   6641
         _ExtentY        =   4022
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox pic临床意义 
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   8265
      ScaleHeight     =   2280
      ScaleWidth      =   3900
      TabIndex        =   31
      Top             =   3735
      Width           =   3900
      Begin VB.TextBox txt参考 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1950
         Left            =   315
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   345
         Width           =   3600
      End
   End
   Begin VB.PictureBox pic诊断 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   8115
      ScaleHeight     =   1185
      ScaleWidth      =   3900
      TabIndex        =   27
      Top             =   1545
      Width           =   3900
      Begin VB.TextBox txt诊断 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   105
         Width           =   4020
      End
   End
   Begin VB.PictureBox pic备注 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   8145
      ScaleHeight     =   1185
      ScaleWidth      =   3900
      TabIndex        =   25
      Top             =   195
      Width           =   3900
      Begin VB.TextBox txt备注 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   120
         Width           =   4020
      End
   End
   Begin VB.PictureBox picRpt 
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   60
      ScaleHeight     =   4590
      ScaleWidth      =   8010
      TabIndex        =   12
      Top             =   165
      Width           =   8010
      Begin VB.CheckBox chk诊断 
         Appearance      =   0  'Flat
         Caption         =   "诊断"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4050
         TabIndex        =   30
         Top             =   30
         Width           =   675
      End
      Begin VB.CheckBox chk备注 
         Appearance      =   0  'Flat
         Caption         =   "备注"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3380
         TabIndex        =   29
         Top             =   30
         Width           =   675
      End
      Begin VB.CheckBox chkChina 
         Appearance      =   0  'Flat
         Caption         =   "中文"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   15
         TabIndex        =   17
         Top             =   30
         Width           =   690
      End
      Begin VB.CheckBox chkMB 
         Appearance      =   0  'Flat
         Caption         =   "酶标"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2725
         TabIndex        =   16
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkReferrence 
         Appearance      =   0  'Flat
         Caption         =   "参考"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2070
         TabIndex        =   15
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkUnit 
         Appearance      =   0  'Flat
         Caption         =   "单位"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1415
         TabIndex        =   14
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkSign 
         Appearance      =   0  'Flat
         Caption         =   "标志"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   700
         TabIndex        =   13
         Top             =   30
         Width           =   720
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   4095
         Left            =   0
         TabIndex        =   24
         Top             =   315
         Width           =   7920
         _cx             =   13970
         _cy             =   7223
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483634
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "警示"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   6540
         TabIndex        =   23
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl警示 
         BackColor       =   &H000040C0&
         Height          =   210
         Left            =   6210
         TabIndex        =   22
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "偏高"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5760
         TabIndex        =   21
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl偏高 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Height          =   210
         Left            =   5430
         TabIndex        =   20
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "偏低"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5010
         TabIndex        =   19
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl偏低 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Height          =   210
         Left            =   4650
         TabIndex        =   18
         Top             =   45
         Width           =   285
      End
   End
   Begin VB.PictureBox picChart 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   75
      ScaleHeight     =   2565
      ScaleWidth      =   9600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4950
      Width           =   9600
      Begin VB.CommandButton cmdRefersh 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6540
         Picture         =   "frmLisRptGeneral.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   465
      End
      Begin VB.TextBox txt次数 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6135
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "3"
         Top             =   90
         Width           =   330
      End
      Begin VB.TextBox txt天数 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4545
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "10"
         Top             =   90
         Width           =   375
      End
      Begin VB.OptionButton opt内容 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "结果值(&2)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   4
         Top             =   75
         Width           =   1125
      End
      Begin VB.OptionButton opt内容 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "变异率(&1)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt内容 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "糖耐量(&3)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2250
         TabIndex        =   2
         Top             =   75
         Width           =   1125
      End
      Begin C1Chart2D8.Chart2D chtThis 
         Height          =   1965
         Left            =   60
         TabIndex        =   5
         Top             =   345
         Width           =   8415
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   14843
         _ExtentY        =   3466
         _StockProps     =   0
         ControlProperties=   "frmLisRptGeneral.frx":685E
      End
      Begin VB.Label lbl项目 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称:RBC"
         Height          =   180
         Left            =   7050
         TabIndex        =   10
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lbl次数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大跟踪次数:"
         Height          =   180
         Left            =   4965
         TabIndex        =   9
         Top             =   90
         Width           =   1170
      End
      Begin VB.Label lbl天数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大跟踪天数:"
         Height          =   180
         Left            =   3390
         TabIndex        =   8
         Top             =   90
         Width           =   1170
      End
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7545
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   8115
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLisRptGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    排列序号 = 0: 检验项目: 英文名: 检验结果: 单位: CV: 结果标志: 结果参考:  OD: CUTOFF: COV: 小数: 检验项目ID: 标本ID: 结果类型 ': 酶标板ID: 变异报警: 变异警示 '结果范围: 固定项目: 警戒上限: 警戒下限: 诊疗项目id:  标本id:
End Enum
Private Const mColCount = 15
Private mstrEndTime As String    '本次检验日期
Private mintIdentMode As Integer    '历史比较病人识别方式
Private mlng医嘱ID As Long
Public mlngMod As Long '调用模块
Private mrsVsf As ADODB.Recordset

Private Sub chkChina_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("查看中文", Me.chkChina.value, glngSys, mlngMod)
End Sub

Private Sub chkMB_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("查看酶标", Me.chkMB.value, glngSys, mlngMod)
End Sub

Private Sub chkReferrence_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("查看参考", Me.chkReferrence.value, glngSys, mlngMod)
End Sub

Private Sub chkSign_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("查看标志", Me.chkSign.value, glngSys, mlngMod)
End Sub

Private Sub chkUnit_Click()
    Call Check_ColWidth
    Call zlDatabase.SetPara("查看单位", Me.chkUnit.value, glngSys, mlngMod)
End Sub

Private Sub chk备注_Click()
    If chk备注.value = 1 Then
        dkpMan.ShowPane 3
    Else
        dkpMan.FindPane(3).Close
    End If
    Call zlDatabase.SetPara("显示备注", Me.chk备注.value, glngSys, mlngMod)
End Sub

Private Sub chk诊断_Click()
    If chk诊断.value = 1 Then
        dkpMan.ShowPane 4
    Else
        dkpMan.FindPane(4).Close
    End If
    Call zlDatabase.SetPara("显示诊断", Me.chk诊断.value, glngSys, mlngMod)
End Sub

Private Sub cmdRefersh_Click()
    Call vsf_RowColChange
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionCollapsing Or Action = PaneActionCollapsed Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1: Item.Handle = Me.picRpt.Hwnd
'    Case 2: Item.Handle = Me.pic参考.hWnd
    Case 3: Item.Handle = Me.pic备注.Hwnd
    Case 4: Item.Handle = Me.pic诊断.Hwnd
    Case 5: Item.Handle = Me.picTab.Hwnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
     Bottom = Me.sbrInfo.Height
End Sub

Private Sub dkpMan_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    'Call RefVsf
End Sub

Private Sub Form_Load()
    '窗格划分
    '-----------------------------------------------------
    Dim panThis As Pane, pan2 As Pane, pan3 As Pane, pan4 As Pane, Pan5 As Pane
    Set panThis = dkpMan.CreatePane(1, 600, 400, DockTopOf, Nothing)
    panThis.Title = "检验报告"
    panThis.Options = PaneNoCaption
    
    Set pan3 = dkpMan.CreatePane(3, 200, 400, DockRightOf, panThis)
    pan3.Title = "检验备注"
    
    Set pan4 = dkpMan.CreatePane(4, 200, 400, DockBottomOf, pan3)
    pan4.Title = "诊断信息"
    
    Set panThis = dkpMan.CreatePane(5, 200, 300, DockBottomOf, Nothing)
    panThis.Title = "历史对比图"
    panThis.Options = PaneNoCaption
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
'    Set pan2 = dkpMan.CreatePane(2, 200, 400, DockRightOf, panThis)
'    pan2.Title = "诊断参考"
    
    Call initVsf
    Call IntiTab
    
    chkChina.value = Val(zlDatabase.GetPara("查看中文", glngSys, mlngMod, 1))
    chkSign.value = Val(zlDatabase.GetPara("查看标志", glngSys, mlngMod, 1))
    chkUnit.value = Val(zlDatabase.GetPara("查看单位", glngSys, mlngMod, 1))
    chkReferrence.value = Val(zlDatabase.GetPara("查看参考", glngSys, mlngMod, 1))
    chkMB.value = Val(zlDatabase.GetPara("查看酶标", glngSys, mlngMod, 1))
    chk备注.value = Val(zlDatabase.GetPara("查看备注", glngSys, mlngMod, 1))
    chk诊断.value = Val(zlDatabase.GetPara("查看诊断", glngSys, mlngMod, 1))
    
    '22539，判断备注及诊断是否保存
    If chk备注.value = 1 Then
        dkpMan.ShowPane 3
    Else
        dkpMan.FindPane(3).Close
    End If
    If chk诊断.value = 1 Then
        dkpMan.ShowPane 4
    Else
        dkpMan.FindPane(4).Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call zlDatabase.SetPara("查看标志", Me.chkSign.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("查看单位", Me.chkUnit.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("查看参考", Me.chkReferrence.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("查看酶标", Me.chkMB.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("查看中文", Me.chkChina.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("查看备注", Me.chk备注.value, glngSys, mlngMod)
    Call zlDatabase.SetPara("查看诊断", Me.chk诊断.value, glngSys, mlngMod)
    mlng医嘱ID = 0
    Set mrsVsf = Nothing
End Sub

Private Sub opt内容_Click(Index As Integer)
    Call vsf_RowColChange
End Sub

Private Sub picChart_Resize()
    err = 0: On Error Resume Next
    With Me.chtThis
        
        .Left = 0
        .Width = Me.picChart.ScaleWidth
        .Height = Me.picChart.ScaleHeight - .Top
        
    End With
    
'    chk参考.Left = Me.picChart.ScaleWidth - chk参考.Width
'    If chk参考.Value = 1 Then
'        Me.txt参考.Top = Me.chtThis.Top
'        Me.txt参考.Left = Me.chtThis.Left + Me.chtThis.Width + 30
'        Me.txt参考.Width = Me.picChart.ScaleWidth - Me.chtThis.Width - 30
'        Me.txt参考.Height = Me.chtThis.Height
'    End If
End Sub

Private Sub picRpt_Resize()
    On Error Resume Next
    With vsf
        .Top = chkChina.Top + chkChina.Height + 10
        .Left = 10
        .Width = picRpt.ScaleWidth - 20
        .Height = picRpt.Height - .Top - 10
    End With
    Call RefVsf
End Sub

Private Sub picTab_Resize()
    With Me.TabThis
        .Top = 0
        .Left = 0
        .Width = Me.picTab.ScaleWidth
        .Height = Me.picTab.ScaleHeight
    End With
End Sub

Private Sub pic备注_Resize()
    With Me.txt备注
    .Left = 0
    .Top = 0
    .Width = Me.pic备注.ScaleWidth
    .Height = Me.pic备注.ScaleHeight
    End With
End Sub

Private Sub pic临床意义_Resize()
    With Me.txt参考
        .Left = 0
        .Top = 0
        .Width = Me.pic临床意义.ScaleWidth
        .Height = Me.pic临床意义.ScaleHeight
    End With
End Sub

Private Sub pic诊断_Resize()
    With Me.txt诊断
        .Left = 0
        .Top = 0
        .Width = Me.pic诊断.ScaleWidth
        .Height = Me.pic诊断.ScaleHeight
    End With
End Sub

Private Sub txt次数_GotFocus()
    Me.txt次数.SelStart = 0: Me.txt次数.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt次数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt天数_GotFocus()
    Me.txt天数.SelStart = 0: Me.txt天数.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt天数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub IntiTab()


    On Error Resume Next

    With Me.TabThis
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearanceExcel
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        .PaintManager.ClientFrame = xtpTabFrameSingleLine
'        .PaintManager.Position = xtpTabPositionBottom
        .InsertItem(0, "图形内容", picChart.Hwnd, conMenu_Tool_Monitor).Tag = "图形内容"
        .InsertItem(1, "临床意义", pic临床意义.Hwnd, conMenu_View_ToolBar_Text).Tag = "临床意义"
        
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        
    End With
End Sub

Public Sub zlRefresh(ByVal lng医嘱ID As Long)
    '显示检验结果,按医嘱ID显示,因有一并核收的情况。
    
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errHandle
    
    vsf.Rows = 1: vsf.Rows = 2
    Me.txt诊断.Text = ""
    strSql = "Select 检验备注,检验人,检验时间,审核人,审核时间 From 检验标本记录 where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng医嘱ID)
    If rsTmp.EOF Then
        Me.txt备注.Text = ""
        
        With sbrInfo
            .Panels(1).Text = "报告人："
            .Panels(2).Text = "报告时间："
            .Panels(3).Text = "审核人："
            .Panels(4).Text = "审核时间："
        End With
        
        Exit Sub
    Else
        Me.txt备注.Text = Trim("" & rsTmp!检验备注)
        
        With sbrInfo
            .Panels(1).Text = "报告人：" & rsTmp!检验人
            .Panels(2).Text = "报告时间：" & IIF(IsNull(rsTmp("检验时间")), "", Format(rsTmp("检验时间"), "yyyy-MM-dd hh:mm"))
            .Panels(3).Text = "审核人：" & rsTmp!审核人
            .Panels(4).Text = "审核时间：" & IIF(IsNull(rsTmp("审核时间")), "", Format(rsTmp("审核时间"), "yyyy-MM-dd hh:mm"))
        End With
    End If
    
    strSql = "Select /*+ rule */" & vbNewLine & _
            "Distinct A.标本id, A.诊疗项目id, A.编码, A.排列序号, A.固定项目, A.ID, A.检验项目,A.缩写 as 英文名, " & vbNewLine & _
            "         A.Cv, Decode(A.本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**', A.本次结果) As 本次结果," & vbNewLine & _
            "         Rownum As 序号, A.标志, A.仪器id, A.标本类别, A.核收时间, A.标本序号, A.标本号显示," & vbNewLine & _
            "         A.检验备注, A.姓名, A.性别, A.年龄, A.门诊号, A.住院号, A.当前床号, A.主页id, A.结果范围," & vbNewLine & _
            "         Nvl(G.小数位数, 2) As 小数, A.警戒上限, A.警戒下限, A.单位," & vbNewLine & _
            "         a.结果参考 As 参考, A.Od," & vbNewLine & _
            "         A.Cutoff, A.Cov, A.酶标板id, A.变异报警, A.变异警示, A.结果类型,A.结果参考" & vbNewLine & _
            "From (Select A.ID As 标本id, B.诊疗项目id, lpad(Decode(D.排列序号, Null, Nvl(H.编码, C.编码), D.排列序号),4,'0') As 编码," & vbNewLine & _
            "              Nvl(B.排列序号, 9999) As 排列序号, Decode(B.诊疗项目id, Null, 0, 1) As 固定项目, B.检验项目id As ID," & vbNewLine & _
            "              C.中文名 || Decode(D.缩写, Null, '', '(' || D.缩写 || ')') As 检验项目, D.缩写,B.原始结果, '' As 上次结果," & vbNewLine & _
            "              '' As 上次时间, '' As Cv, B.检验结果 As 本次结果, D.计算公式, D.结果类型," & vbNewLine & _
            "              Decode(B.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
            "              Nvl(A.仪器id, -1) As 仪器id, Nvl(A.标本类别, 0) As 标本类别, A.核收时间, A.标本序号," & vbNewLine & _
            "              Decode(A.仪器id, Null," & vbNewLine & _
            "                      To_Char(Trunc(A.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(A.标本序号, 10000), '0000')," & vbNewLine & _
            "                      A.标本序号) As 标本号显示, A.检验备注, A.姓名, A.性别, A.年龄, A.标本类型, A.出生日期, A.门诊号," & vbNewLine & _
            "              A.住院号, A.床号 As 当前床号, A.主页id, D.结果范围, D.警戒上限, D.警戒下限, D.单位, B.Od, B.Cutoff," & vbNewLine & _
            "              B.Sco As Cov, B.酶标板id, D.变异报警率 As 变异报警, D.变异警示率 As 变异警示,B.结果参考" & vbNewLine & _
            "       From 检验标本记录 A, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 诊疗项目目录 H" & vbNewLine & _
            "       Where A.ID = B.检验标本id And B.检验项目id = C.ID And C.ID = D.诊治项目id And B.诊疗项目id = H.ID(+) And" & vbNewLine & _
            "             B.记录类型 = A.报告结果 And A.医嘱ID = [1]"
    strSql = strSql & "       Union All" & vbNewLine & _
            "       Select A.ID As 标本id, B.诊疗项目id, lpad(Decode(D.排列序号, Null, Nvl(H.编码, C.编码), D.排列序号),4,'0') As 编码," & vbNewLine & _
            "              Nvl(B.排列序号, 9999) As 排列序号, Decode(B.诊疗项目id, Null, 0, 1) As 固定项目,B.检验项目id As ID," & vbNewLine & _
            "              C.中文名 || Decode(D.缩写, Null, '', '(' || D.缩写 || ')') As 检验项目,D.缩写,B.原始结果, '' As 上次结果," & vbNewLine & _
            "              '' As 上次时间, '' As Cv, B.检验结果 As 本次结果, D.计算公式, D.结果类型," & vbNewLine & _
            "              Decode(B.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
            "              Nvl(A.仪器id, -1) As 仪器id, Nvl(A.标本类别, 0) As 标本类别, A.核收时间, A.标本序号," & vbNewLine & _
            "              Decode(A.仪器id, Null," & vbNewLine & _
            "                      To_Char(Trunc(A.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(A.标本序号, 10000), '0000')," & vbNewLine & _
            "                      A.标本序号) As 标本号显示, A.检验备注, A.姓名, A.性别, A.年龄, A.标本类型, A.出生日期, A.门诊号," & vbNewLine & _
            "              A.住院号, A.床号 As 当前床号, A.主页id, D.结果范围, D.警戒上限, D.警戒下限, D.单位, B.Od, B.Cutoff," & vbNewLine & _
            "              B.Sco As Cov, B.酶标板id, D.变异报警率 As 变异报警, D.变异警示率 As 变异警示,B.结果参考" & vbNewLine & _
            "       From 检验标本记录 A,检验标本记录 E, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 检验仪器项目 G, 诊疗项目目录 H" & vbNewLine & _
            "       Where A.ID = B.检验标本id And B.检验项目id = C.ID And C.ID = D.诊治项目id And B.诊疗项目id = H.ID(+) And" & vbNewLine & _
            "             B.记录类型 = A.报告结果 And E.ID=A.合并id  And E.医嘱ID= [1]) A, 检验仪器项目 G" & vbNewLine & _
            "Where A.仪器id = G.仪器id(+) And A.ID = G.项目id(+)" & vbNewLine & _
            "Order By A.编码, A.排列序号"
    
    Set mrsVsf = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng医嘱ID)
    
    Call RefVsf
    '显示诊断信息
    Dim strTmp As String
    strSql = "Select distinct b.医嘱id, b.项目, b.排列, b.内容" & vbNewLine & _
                "From 检验标本记录 a, 病人医嘱附件 b" & vbNewLine & _
                "Where a.医嘱id = b.医嘱id and a.医嘱ID = [1] " & vbNewLine & _
                "Order By 医嘱id, 排列"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng医嘱ID)
    Do Until rsTmp.EOF
        strTmp = strTmp & Trim("" & rsTmp("项目")) & ":" & Replace(Trim("" & rsTmp("内容")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rsTmp.MoveNext
    Loop
    Me.txt诊断.Text = strTmp
    
    Call RefChartData(lng医嘱ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefVsf()
    Dim lngRow As Long, lngCol As Long
    Dim bln分栏 As Boolean
    Call initVsf
    If mrsVsf Is Nothing Then Exit Sub
    
    lngRow = vsf.FixedRows
    If mrsVsf.RecordCount > 0 Then mrsVsf.MoveFirst
    Do Until mrsVsf.EOF
        With vsf
            Dim lngAdd As Long
            lngAdd = 1
            'If .ScrollBars >= flexScrollBarVertical Then lngAdd = 2
            If (Split(Format(.ClientHeight / .RowHeightMin, "0.0000"), ".")(0)) >= lngRow + 2 And bln分栏 = False Then
                lngCol = 0
            Else
                If lngRow > 5 Then '至少5行
                    If bln分栏 = False Then
                        bln分栏 = True
                        Call Add_Column(lngCol, lngRow)
                    End If
                End If
            End If
            If (lngCol = 0 And lngRow >= .Rows) Or (lngCol > 0 And lngRow >= .Rows - 1) Then
                Call Add_Column(lngCol, lngRow)
            End If

            .TextMatrix(lngRow, mCol.排列序号 + lngCol * mColCount) = mrsVsf.Bookmark  'Trim("" & mrsVsf!排列序号)
            .TextMatrix(lngRow, mCol.检验项目 + lngCol * mColCount) = Trim("" & mrsVsf!检验项目)
            .TextMatrix(lngRow, mCol.英文名 + lngCol * mColCount) = Trim("" & mrsVsf!英文名)
            .TextMatrix(lngRow, mCol.检验结果 + lngCol * mColCount) = Trim("" & mrsVsf!本次结果)
            .TextMatrix(lngRow, mCol.单位 + lngCol * mColCount) = Trim("" & mrsVsf!单位)
            .TextMatrix(lngRow, mCol.CV + lngCol * mColCount) = Trim("" & mrsVsf!CV)
            .TextMatrix(lngRow, mCol.结果标志 + lngCol * mColCount) = Trim("" & mrsVsf!标志)
            .TextMatrix(lngRow, mCol.结果参考 + lngCol * mColCount) = IIF(Trim("" & mrsVsf!参考) = "", Trim("" & mrsVsf!结果参考), Trim("" & mrsVsf!参考))
            .TextMatrix(lngRow, mCol.OD + lngCol * mColCount) = Trim("" & mrsVsf!OD)
            .TextMatrix(lngRow, mCol.CUTOFF + lngCol * mColCount) = Trim("" & mrsVsf!CUTOFF)
            .TextMatrix(lngRow, mCol.COV + lngCol * mColCount) = Trim("" & mrsVsf!COV)
            .TextMatrix(lngRow, mCol.小数 + lngCol * mColCount) = Trim("" & mrsVsf!小数)
            .TextMatrix(lngRow, mCol.检验项目ID + lngCol * mColCount) = Trim("" & mrsVsf!ID)
            .TextMatrix(lngRow, mCol.标本ID + lngCol * mColCount) = Trim("" & mrsVsf!标本ID)
            .TextMatrix(lngRow, mCol.结果类型 + lngCol * mColCount) = Trim("" & mrsVsf!结果类型)
            
            If lngCol = 0 Then
                .Rows = .Rows + 1
            End If
            lngRow = lngRow + 1
        End With
        mrsVsf.MoveNext
    Loop
    Call Check_ColWidth
    
    vsf.Rows = vsf.Rows - 1
End Sub
Private Sub RefChartData(ByVal lng医嘱ID As Long)
    '更新显示 最大跟踪次数
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mlng医嘱ID = lng医嘱ID Then Exit Sub
    mlng医嘱ID = lng医嘱ID
    strSql = "Select /* +rule */" & vbNewLine & _
        " Nvl(L.检验时间, Sysdate) As 检验时间, Nvl(Max(跟踪天数), 0) As 天数" & vbNewLine & _
        "From 检验项目选项 O, 检验报告项目 X, 检验普通结果 R, 检验标本记录 L" & vbNewLine & _
        "Where O.诊疗项目id(+) = X.诊疗项目id And X.报告项目id = R.检验项目id And R.检验标本id = L.ID And L.医嘱ID = [1]" & vbNewLine & _
        "Group By Nvl(L.检验时间, Sysdate)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng医嘱ID)
    Me.txt天数.Text = 30
    mstrEndTime = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    Do Until rsTmp.EOF
         Me.txt天数.Text = rsTmp!天数
         mstrEndTime = Format(rsTmp!检验时间, "yyyy-MM-dd hh:mm:ss")
        rsTmp.MoveNext
    Loop
    If Val(Me.txt天数.Text) <= 0 Then Me.txt天数.Text = 30
    If Val(Me.txt次数.Text) <= 0 Then Me.txt次数 = 3
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Add_Column(ByRef lngCol As Long, ByRef lngRow As Long)
    '添加分栏所需的列
    With vsf
        lngCol = lngCol + 1
        lngRow = vsf.FixedRows
        .Cols = .Cols + mColCount
        .TextMatrix(0, mCol.排列序号 + lngCol * mColCount) = "": .ColWidth(mCol.排列序号 + lngCol * mColCount) = 300: .ColAlignment(mCol.排列序号 + lngCol * mColCount) = flexAlignRightCenter
        .TextMatrix(0, mCol.检验项目 + lngCol * mColCount) = "检验项目": .ColWidth(mCol.检验项目 + lngCol * mColCount) = 2100: .ColAlignment(mCol.检验项目 + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.英文名 + lngCol * mColCount) = "检验项目": .ColWidth(mCol.英文名 + lngCol * mColCount) = 1000: .ColAlignment(mCol.英文名 + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.检验结果 + lngCol * mColCount) = "检验结果": .ColWidth(mCol.检验结果 + lngCol * mColCount) = 1200: .ColAlignment(mCol.检验结果 + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.单位 + lngCol * mColCount) = "单位": .ColWidth(mCol.单位 + lngCol * mColCount) = 1000: .ColAlignment(mCol.单位 + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CV + lngCol * mColCount) = "CV": .ColWidth(mCol.CV + lngCol * mColCount) = 0: .ColAlignment(mCol.CV + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.CV + lngCol * mColCount) = True
        .TextMatrix(0, mCol.结果标志 + lngCol * mColCount) = "标志": .ColWidth(mCol.结果标志 + lngCol * mColCount) = 450: .ColAlignment(mCol.结果标志 + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.结果参考 + lngCol * mColCount) = "参考": .ColWidth(mCol.结果参考 + lngCol * mColCount) = 1300: .ColAlignment(mCol.结果参考 + lngCol * mColCount) = flexAlignLeftCenter
        
        .TextMatrix(0, mCol.OD + lngCol * mColCount) = "OD": .ColWidth(mCol.OD + lngCol * mColCount) = 700: .ColAlignment(mCol.OD + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CUTOFF + lngCol * mColCount) = "CUTOFF": .ColWidth(mCol.CUTOFF + lngCol * mColCount) = 700: .ColAlignment(mCol.CUTOFF + lngCol * mColCount) = flexAlignLeftCenter
        .TextMatrix(0, mCol.COV + lngCol * mColCount) = "COV": .ColWidth(mCol.COV + lngCol * mColCount) = 700: .ColAlignment(mCol.COV + lngCol * mColCount) = flexAlignLeftCenter
        
        .TextMatrix(0, mCol.小数 + lngCol * mColCount) = "小数": .ColWidth(mCol.小数 + lngCol * mColCount) = 0: .ColAlignment(mCol.小数 + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.小数 + lngCol * mColCount) = True
        .TextMatrix(0, mCol.检验项目ID + lngCol * mColCount) = "检验项目id": .ColWidth(mCol.检验项目ID + lngCol * mColCount) = 0: .ColAlignment(mCol.检验项目ID + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.检验项目ID + lngCol * mColCount) = True
        .TextMatrix(0, mCol.标本ID + lngCol * mColCount) = "标本ID": .ColWidth(mCol.标本ID + lngCol * mColCount) = 0: .ColAlignment(mCol.标本ID + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.标本ID + lngCol * mColCount) = True
        .TextMatrix(0, mCol.结果类型 + lngCol * mColCount) = "结果类型": .ColWidth(mCol.结果类型 + lngCol * mColCount) = 0: .ColAlignment(mCol.结果类型 + lngCol * mColCount) = flexAlignLeftCenter
        .ColHidden(mCol.结果类型 + lngCol * mColCount) = True
        
        
    End With
End Sub
Private Sub initVsf()
    '初始化表格
    With vsf
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 2: .FixedRows = 1
        .Cols = mColCount: .FixedCols = 0
        
        .TextMatrix(0, mCol.排列序号) = "": .ColWidth(mCol.排列序号) = 300: .ColAlignment(mCol.排列序号) = flexAlignRightCenter
        .TextMatrix(0, mCol.检验项目) = "检验项目": .ColWidth(mCol.检验项目) = 2100: .ColAlignment(mCol.检验项目) = flexAlignLeftCenter
        .TextMatrix(0, mCol.英文名) = "检验项目": .ColWidth(mCol.检验项目) = 1000: .ColAlignment(mCol.英文名) = flexAlignLeftCenter
        
        .TextMatrix(0, mCol.检验结果) = "检验结果": .ColWidth(mCol.检验结果) = 1200: .ColAlignment(mCol.检验结果) = flexAlignLeftCenter
        .TextMatrix(0, mCol.单位) = "单位": .ColWidth(mCol.单位) = 1000: .ColAlignment(mCol.单位) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CV) = "CV": .ColWidth(mCol.CV) = 0: .ColAlignment(mCol.CV) = flexAlignLeftCenter
        .ColHidden(mCol.CV) = True
        .TextMatrix(0, mCol.结果标志) = "标志": .ColWidth(mCol.结果标志) = 450: .ColAlignment(mCol.结果标志) = flexAlignLeftCenter
        .TextMatrix(0, mCol.结果参考) = "参考": .ColWidth(mCol.结果参考) = 1300: .ColAlignment(mCol.结果参考) = flexAlignLeftCenter


        .TextMatrix(0, mCol.OD) = "OD": .ColWidth(mCol.OD) = 700: .ColAlignment(mCol.OD) = flexAlignLeftCenter
        .TextMatrix(0, mCol.CUTOFF) = "CUTOFF": .ColWidth(mCol.CUTOFF) = 700: .ColAlignment(mCol.CUTOFF) = flexAlignLeftCenter
        .TextMatrix(0, mCol.COV) = "COV": .ColWidth(mCol.COV) = 700: .ColAlignment(mCol.COV) = flexAlignLeftCenter
        .TextMatrix(0, mCol.小数) = "小数": .ColWidth(mCol.小数) = 0: .ColAlignment(mCol.小数) = flexAlignLeftCenter
        .ColHidden(mCol.小数) = True
        .TextMatrix(0, mCol.检验项目ID) = "检验项目ID": .ColWidth(mCol.检验项目ID) = 0: .ColAlignment(mCol.检验项目ID) = flexAlignLeftCenter
        .ColHidden(mCol.检验项目ID) = True
        .TextMatrix(0, mCol.标本ID) = "标本ID": .ColWidth(mCol.标本ID) = 0: .ColAlignment(mCol.标本ID) = flexAlignLeftCenter
        .ColHidden(mCol.标本ID) = True
        .TextMatrix(0, mCol.结果类型) = "结果类型": .ColWidth(mCol.结果类型) = 0: .ColAlignment(mCol.结果类型) = flexAlignLeftCenter
        .ColHidden(mCol.结果类型) = True
        
        Call Check_ColWidth
    End With
End Sub

Private Sub Check_ColWidth()
    '根据控件状态，调整列宽
    
    Dim lngCol As Long, lngLoop As Long, lngRow As Long
    Dim lngColor As Long, lngForeColor As Long, str标志 As String
    With vsf
        lngCol = (.Cols / mColCount)
        For lngLoop = 0 To lngCol - 1
            '序号列的颜色
            .Cell(flexcpBackColor, 1, mCol.排列序号 + lngLoop * mColCount, vsf.Rows - 1, mCol.排列序号 + lngLoop * mColCount) = vsf.BackColorFixed
            
            '列宽
            .ColWidth(mCol.检验项目 + lngLoop * mColCount) = IIF(chkChina.value = 0, 0, 2100)
            .ColWidth(mCol.英文名 + lngLoop * mColCount) = IIF(chkChina.value = 0, 1000, 0)
            .ColWidth(mCol.结果标志 + lngLoop * mColCount) = IIF(chkSign.value = 0, 0, 450)
            .ColWidth(mCol.单位 + lngLoop * mColCount) = IIF(chkUnit.value = 0, 0, 1000)
            .ColWidth(mCol.结果参考 + lngLoop * mColCount) = IIF(chkReferrence.value = 0, 0, 1300)
            
            .ColWidth(mCol.OD + lngLoop * mColCount) = IIF(chkMB.value = 0, 0, 700)
            .ColWidth(mCol.CUTOFF + lngLoop * mColCount) = IIF(chkMB.value = 0, 0, 700)
            .ColWidth(mCol.COV + lngLoop * mColCount) = IIF(chkMB.value = 0, 0, 700)
            
            For lngRow = .FixedRows To .Rows - 1
                '单元格格式
                If IsNumeric("-" & .TextMatrix(lngRow, mCol.检验结果 + lngLoop * mColCount)) Then
                    .TextMatrix(lngRow, mCol.检验结果 + lngLoop * mColCount) = Format(.TextMatrix(lngRow, mCol.检验结果 + lngLoop * mColCount), _
                        IIF(Val(.TextMatrix(lngRow, mCol.小数 + lngLoop * mColCount)) = 0, "#0", "0." & String(Val(.TextMatrix(lngRow, mCol.小数 + lngLoop * mColCount)), "0")))
                End If
                '颜色
                lngColor = .BackColor
                lngForeColor = .ForeColor
                str标志 = Trim(.TextMatrix(lngRow, mCol.结果标志 + lngLoop * mColCount))
                If InStr("↓", str标志) > 0 And str标志 <> "" Then     '2
                    lngColor = lbl偏低.BackColor
                    lngForeColor = lbl偏低.ForeColor
                ElseIf InStr("↑,异常", str标志) > 0 And str标志 <> "" Then '3,异常
                    lngColor = lbl偏高.BackColor
                    lngForeColor = lbl偏高.ForeColor
                ElseIf InStr("↑↑,↓↓", str标志) > 0 And str标志 <> "" Then '5,6
                    lngColor = lbl警示.BackColor
                    lngForeColor = lbl警示.ForeColor
                End If
                .Cell(flexcpBackColor, lngRow, mCol.检验结果 + lngLoop * mColCount, lngRow, mCol.检验结果 + lngLoop * mColCount) = lngColor
                .Cell(flexcpForeColor, lngRow, mCol.检验结果 + lngLoop * mColCount, lngRow, mCol.检验结果 + lngLoop * mColCount) = lngForeColor
            Next
        Next
    End With
End Sub

Private Function get_Column(ByVal lngCol As Long) As Long
    '得到指定列是第几个分栏
    Dim strTmp As String
    strTmp = CStr(Format(lngCol / mColCount, "0.00000"))
    If InStr(strTmp, ".") > 0 Then
        get_Column = Val(Mid(strTmp, 1, InStr(strTmp, ".")))
    Else
        get_Column = Val(strTmp)
    End If
End Function

Private Sub vsf_RowColChange()
    Dim lng项目id As Long, str项目 As String, str结果类型 As String, lng标本id As Long
    Dim lngCol As Long
    
    lngCol = get_Column(vsf.Col)
    
    lng项目id = Val(vsf.TextMatrix(vsf.Row, mCol.检验项目ID + lngCol * mColCount))
    lng标本id = Val(vsf.TextMatrix(vsf.Row, mCol.标本ID + lngCol * mColCount))
    If chkChina.value Then
        str项目 = vsf.TextMatrix(vsf.Row, mCol.检验项目 + lngCol * mColCount)
    Else
        str项目 = vsf.TextMatrix(vsf.Row, mCol.英文名 + lngCol * mColCount)
    End If
    Me.lbl项目.Caption = "项目：" & str项目
    Me.lbl项目.Left = Me.cmdRefersh.Left + Me.cmdRefersh.Width + 45
    
    str结果类型 = vsf.TextMatrix(vsf.Row, mCol.结果类型 + lngCol * mColCount)
    If lng项目id <> 0 Then
        Call RefChart(lng标本id, lng项目id, str结果类型)
        
    End If
End Sub


Private Sub RefChart(ByVal lng标本id As Long, ByVal lng项目id As Long, ByVal str结果类型 As String)
    '画分析图
    Dim aryX() As Variant, aryY() As Variant
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngDates As Long, lngCount As Long, intLoop As Integer, dblAvg As Double, lng次数 As Long
    Dim strMaxValue As String, strMinValue As String
    Dim dblCurCV As Double, dbl本次结果 As Double, dbl变异报警率 As Double
    On Error GoTo errHandle
    '将序列数字设置为0，清除图形显示
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    
    
    '临床意义
    txt参考.Text = ""
    strSql = "Select 临床意义 From 检验项目 Where 诊治项目id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng项目id)
    Do Until rsTmp.EOF
        txt参考.Text = Trim("" & rsTmp.Fields("临床意义"))
        rsTmp.MoveNext
    Loop
    
    
    If str结果类型 = "2" Or str结果类型 = "3" Then
       Me.chtThis.IsBatched = False
       Exit Sub
    End If
    
    '设置图形的基本形态
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot  '折线
        .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
        With .Data
            .Layout = oc2dDataArray
            .NumSeries = 1
            .NumPoints(1) = 4
        End With
    End With
    With Me.chtThis.ChartArea
        .Axes("X").MajorGrid.Spacing.IsDefault = True
        .Axes("Y").MajorGrid.Spacing.IsDefault = True
        .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示
    End With
    
    If Me.opt内容(0).value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "变异率"
    ElseIf Me.opt内容(1).value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "结果值"
    Else
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "糖耐量"
    End If
    
    '提数据
    lngDates = Val(Me.txt天数.Text)
    lng次数 = Val(Me.txt次数.Text)
    
    If Me.opt内容(2).value = True Then
        strSql = "Select 编码, 中文名, 英文名, 检验项目id, 检验结果, decode(别名,null,中文名,别名) as 名称 " & vbNewLine & _
                    "From (Select Decode(E.排列序号, Null, D.编码, E.排列序号) As 编码, D.中文名, D.英文名, B.检验项目id, B.检验结果, H.别名" & vbNewLine & _
                    "       From 检验标本记录 A, 检验普通结果 B, 检验仪器项目 C, 诊治所见项目 D, 检验项目 E, 检验报告项目 F, 诊疗项目目录 G," & vbNewLine & _
                    "            (Select 诊疗项目id, 名称 As 别名 From 诊疗项目别名 Where 性质 = 9 And 码类 = 1) H" & vbNewLine & _
                    "       Where A.ID = B.检验标本id And B.检验项目id = C.项目id And B.检验项目id = D.ID And Nvl(C.糖耐量项目, 0) = -1 And A.医嘱ID = [1] And" & vbNewLine & _
                    "             B.检验项目id = E.诊治项目id And B.检验项目id = F.报告项目id And F.诊疗项目id = G.ID And Nvl(G.组合项目, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.诊疗项目id(+)" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(E.排列序号, Null, D.编码, E.排列序号) As 编码, D.中文名, D.英文名, B.检验项目id, B.检验结果, H.别名" & vbNewLine & _
                    "       From 检验标本记录 A, 检验标本记录 I,检验普通结果 B, 检验仪器项目 C, 诊治所见项目 D, 检验项目 E, 检验报告项目 F, 诊疗项目目录 G," & vbNewLine & _
                    "            (Select 诊疗项目id, 名称 As 别名 From 诊疗项目别名 Where 性质 = 9 And 码类 = 1) H" & vbNewLine & _
                    "       Where A.ID = B.检验标本id And B.检验项目id = C.项目id And B.检验项目id = D.ID And Nvl(C.糖耐量项目, 0) = -1 And A.合并id = I.ID And I.医嘱ID= [1] And" & vbNewLine & _
                    "             B.检验项目id = E.诊治项目id And B.检验项目id = F.报告项目id And F.诊疗项目id = G.ID And Nvl(G.组合项目, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.诊疗项目id(+))" & vbNewLine & _
                    "Order By 编码"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng医嘱ID)
        If rsTmp.RecordCount > 0 Then
            ReDim aryX(rsTmp.RecordCount - 1)
            ReDim aryY(rsTmp.RecordCount - 1, 0)
        End If
    Else
        strSql = "Select I.ID, I.名称 As 中文名, V.缩写 As 英文名, I.计算单位 As 单位, L.次数, L.检验时间, L.检验结果, V.变异报警率,V.结果类型 " & vbNewLine & _
                "From (Select L.检验项目id, L.次数, L.检验时间, L.检验结果 " & vbNewLine & _
                "       From (Select M.病人id As 病人id, M.姓名, M.性别, L.ID As 次数, L.检验时间, R.检验项目id, R.检验结果,L.标本类型 " & vbNewLine & _
                "              From 检验标本记录 L, 检验普通结果 R, 病人医嘱记录 M, (select 病人id,姓名,性别 from 检验标本记录 where 医嘱id = [1]) N " & vbNewLine & _
                "              Where M.ID = L.医嘱id And L.ID = R.检验标本id And  " & vbNewLine & _
                "                    L.检验时间 Between [2]  And" & vbNewLine & _
                "                    [3] and " & IIF(mintIdentMode = 0, "L.病人id = N.病人id", "L.姓名 = N.姓名 And L.性别 = N.性别") & ") L," & vbNewLine & _
                "            (Select M.病人id As 病人id, M.姓名, M.性别, L.检验时间, R.检验项目id,L.标本类型 " & vbNewLine & _
                "              From 病人医嘱记录 M, 检验标本记录 L, 检验普通结果 R" & vbNewLine & _
                "              Where M.ID = L.医嘱id And L.ID = R.检验标本id And L.医嘱id = [1]) C" & vbNewLine & _
                "       Where " & IIF(mintIdentMode = 0, "L.病人id = C.病人id", "L.姓名 = C.姓名 And L.性别 = C.性别") & " And L.检验项目id+0 = C.检验项目id " & _
                "       And L.标本类型 = C.标本类型 ) L, 检验项目 V, 检验报告项目 R, 诊疗项目目录 I" & vbNewLine & _
                "Where L.检验项目id=[4] and L.检验项目id = V.诊治项目id And L.检验项目id = R.报告项目id And R.诊疗项目id = I.ID And I.组合项目 <> 1" & vbNewLine & _
                "Order By I.编码, L.检验时间 desc"
                
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng医嘱ID, CDate(Format(mstrEndTime, "yyyy-MM-dd 00:00:00")) - lngDates, _
                                           CDate(Format(mstrEndTime, "yyyy-MM-dd 23:59:59")), lng项目id)
        If rsTmp.RecordCount > 0 Then
            ReDim aryX(rsTmp.RecordCount - 1)
            ReDim aryY(rsTmp.RecordCount - 1, 0)
            If lng次数 >= rsTmp.RecordCount Then
                ReDim aryX(rsTmp.RecordCount - 1)
                ReDim aryY(rsTmp.RecordCount - 1, 0)
            Else
                ReDim aryX(lng次数 - 1)
                ReDim aryY(lng次数 - 1, 0)
            End If
        End If
    End If
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    
    '填充数据
    If rsTmp.RecordCount > 0 Then
        For lngCount = LBound(aryX) To UBound(aryX)
            
            aryX(lngCount) = lngCount
            If Me.opt内容(0).value = True Then  '变异率
                If lng标本id = rsTmp!次数 Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, "本次结果"
                    dbl本次结果 = Val("" & rsTmp!检验结果)
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Format(rsTmp!检验时间, "yyyy-MM-dd HH:mm")
                End If
                
                If Val(Trim("" & rsTmp!检验结果)) = 0 Or dbl本次结果 = 0 Then
                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Else
                    dblCurCV = Format((Val(Trim("" & rsTmp!检验结果)) - dbl本次结果) / dbl本次结果 * 100, "0.00")
                    aryY(lngCount, 0) = dblCurCV
                End If
                'Debug.Print "结果：" & Val(Trim("" & rsTmp!检验结果)) & ", 变异率:" & dblCurCV
                dbl变异报警率 = Val("" & rsTmp!变异报警率)
                
            ElseIf Me.opt内容(1).value = True Then '结果值
                If lng标本id = rsTmp!次数 Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, "本次结果"
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Format(rsTmp!检验时间, "yyyy-MM-dd HH:mm")
                End If
                If Val(Trim("" & rsTmp!检验结果)) = 0 Then
                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Else
                    aryY(lngCount, 0) = Val(Trim("" & rsTmp!检验结果))
                End If
            Else                                '糖耐量
                Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Trim("" & rsTmp("名称"))
                aryY(lngCount, 0) = Val("" & rsTmp("检验结果"))
                
            End If
            
            rsTmp.MoveNext
            
            If Val(strMaxValue) < Abs(Val(aryY(lngCount, 0))) Then
                strMaxValue = Abs(Val(aryY(lngCount, 0)))
            End If
            If Val(strMinValue) > Abs(Val(aryY(lngCount, 0))) Then
                strMinValue = Abs(Val(aryY(lngCount, 0)))
            End If
            
        Next
        
        '变更刷新内部数据
        Me.chtThis.IsBatched = True
        Me.chtThis.ChartGroups(1).Data.NumPoints(1) = UBound(aryX) + 1
        Call Me.chtThis.ChartGroups(1).Data.CopyXVectorIn(1, aryX)
        Call Me.chtThis.ChartGroups(1).Data.CopyYArrayIn(aryY)
        
        If opt内容(0).value = True Then
            Me.chtThis.ChartArea.Axes("Y").Origin = 0
            Me.chtThis.ChartArea.Axes("Y").Min = -1 * Val(strMaxValue)
            Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
        ElseIf opt内容(1).value = True Then
            On Error Resume Next
            For intLoop = 0 To UBound(aryY, 1) - 1
                dblAvg = dblAvg + Val(aryY(intLoop, 0))
            Next
            If dblAvg <> 0 Then
                dblAvg = dblAvg / UBound(aryY, 1)
                Me.chtThis.ChartArea.Axes("Y").Origin = dblAvg
                If (dblAvg - Val(strMinValue)) < (Val(strMaxValue) - dblAvg) Then
                    Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (Val(strMaxValue) - dblAvg))
                    Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (Val(strMaxValue) - dblAvg))
                Else
                    Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (dblAvg - Val(strMinValue)))
                    Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (dblAvg - Val(strMinValue)))
                End If
            End If
        Else
            Me.chtThis.ChartArea.Axes("Y").Origin = 0
            Me.chtThis.ChartArea.Axes("Y").Min = 0
            Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
        End If
    End If
    Me.chtThis.IsBatched = False
    
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


