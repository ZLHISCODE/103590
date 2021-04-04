VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmTurnToWardFeeQuery 
   Caption         =   "转病区费用查询"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTurnToWardFeeQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   4
      Left            =   -450
      ScaleHeight     =   915
      ScaleWidth      =   12015
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6510
      Width           =   12015
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   10290
         TabIndex        =   14
         ToolTipText     =   "热键:Esc"
         Top             =   300
         Width           =   1380
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   8790
         TabIndex        =   13
         ToolTipText     =   "热键：F2"
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   3
      Left            =   870
      ScaleHeight     =   2055
      ScaleWidth      =   3615
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3615
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   1485
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Top             =   450
         Width           =   2565
         _cx             =   4524
         _cy             =   2619
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         Caption         =   "卫材申请取消"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Tag             =   "卫材销账申请取消"
         Top             =   120
         Width           =   1260
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   2
      Left            =   960
      ScaleHeight     =   2055
      ScaleWidth      =   3615
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2910
      Width           =   3615
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   1485
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   540
         Width           =   2565
         _cx             =   4524
         _cy             =   2619
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卫材销账申请"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Tag             =   "卫生材料"
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   900
      ScaleHeight     =   2055
      ScaleWidth      =   6645
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   780
      Width           =   6645
      Begin VB.PictureBox picCboBack 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   3030
         ScaleHeight     =   315
         ScaleWidth      =   2655
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   390
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox cboDate 
            Height          =   360
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   -30
            Width           =   2715
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   1485
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2565
         _cx             =   4524
         _cy             =   2619
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   10
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   0
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
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "变动时间: 2017-06-12"
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   15
         Tag             =   "变动时间: "
         Top             =   120
         Width           =   2400
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12135
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数: 第1次"
         Height          =   240
         Index           =   8
         Left            =   8280
         TabIndex        =   18
         Tag             =   "住院次数: "
         Top             =   180
         Width           =   1800
      End
      Begin VB.Shape shp病人信息 
         BorderColor     =   &H8000000A&
         Height          =   105
         Left            =   60
         Top             =   0
         Width           =   11715
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别: 普通"
         Height          =   240
         Index           =   9
         Left            =   10380
         TabIndex        =   5
         Tag             =   "费别: "
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号: 99999999"
         Height          =   240
         Index           =   7
         Left            =   6090
         TabIndex        =   4
         Tag             =   "住院号: "
         Top             =   180
         Width           =   1920
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄: 26岁3月"
         Height          =   240
         Index           =   6
         Left            =   4020
         TabIndex        =   3
         Tag             =   "年龄: "
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别: 男"
         Height          =   240
         Index           =   5
         Left            =   2940
         TabIndex        =   2
         Tag             =   "性别: "
         Top             =   180
         Width           =   960
      End
      Begin VB.Line lineShow 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   900
         X2              =   2400
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人: 王麻子"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   1
         Tag             =   "病人: "
         Top             =   180
         Width           =   1440
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   660
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTurnToWardFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Private mbyt操作 As Fun_Index
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng变动id As Long
Private mlng原病区id As Long
Private mlng目标病区id As Long
Private mcllSQL As Collection
'----------------------------------------------------------------------------
Private mblnOK As Boolean

Private Enum IndexDef
    Lbl_变动时间 = 1
    Lbl_卫材申请 = 2
    Lbl_卫材取消 = 3
    Lbl_姓名 = 4
    Lbl_性别 = 5
    Lbl_年龄 = 6
    Lbl_住院号 = 7
    Lbl_住院次数 = 8
    Lbl_费别 = 9
    
    Pane_病人信息 = 0
    Pane_变动费用 = 1
    Pane_卫材申请 = 2
    Pane_卫材取消 = 3
    Pane_功能按钮 = 4
End Enum

Private Enum Fun_Index
    Fun_转病区 = 0
    Fun_撤消转病区 = 1
    Fun_转病区申请 = 2
    Fun_历史转出查询 = 3
End Enum

Public Function TurnToWard_Fee_Query(ByVal frmMain As Object, ByVal byt操作 As Byte, _
    ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    Optional ByVal lng变动id As Long, _
    Optional ByVal lng原病区id As Long, Optional lng目标病区id As Long, _
    Optional ByRef cllSql As Collection) As Boolean
     '------------------------------------------------------------------------------------
    '功能:转病区费用查询
    '入参:frmMain-调用的主窗体
    '     byt操作- 0-转病区;1-撤消病区;2-转病区申请;3-历史转出查询
    '     cllSQL - 需要执行的SQL语句，0-转病区/1-撤消病区有效；
    '               0-转病区时，在更改变动记录之后执行；1-撤消病区时，在更改变动记录之前执行
    '   1、int操作=0(转病区)时
    '       lng变动id: 原病区的变动记录的ID
    '       lng原病区id：原病区ID
    '       lng目标病区id:目标病区ID
    '   2、int操作=1(撤消病区)时
    '       lng变动id: 恢复的原病区的变动记录的ID
    '       lng原病区id：被撤消的病区ID
    '       lng目标病区id:恢复的原始病区ID
    '  3. int操作=3(历史转出查询)
    '出参:
    '返回:费用查询时操作员点确认返回true,否则返回False
    '-----------------------------------------------------------------------------------
    mbyt操作 = byt操作
    mlng病人ID = lng病人ID: mlng主页ID = lng主页Id
    mlng变动id = lng变动id
    mlng原病区id = lng原病区id: mlng目标病区id = lng目标病区id
    
    mblnOK = False
    On Error Resume Next
    Me.Show 1, frmMain
    
    If mblnOK Then
        If Not mcllSQL Is Nothing Then
            Set cllSql = mcllSQL
        End If
        TurnToWard_Fee_Query = True
    End If
End Function

Private Sub Form_Load()
    Dim objPane As Pane
    
    Me.Width = 1024 * Screen.TwipsPerPixelX
    Me.Height = 768 * Screen.TwipsPerPixelY
    picCboBack.Visible = (mbyt操作 = Fun_历史转出查询)
    
    If mbyt操作 = Fun_转病区申请 Or mbyt操作 = Fun_转病区 Then
        If Upgrade医嘱执行计价执行状态(mlng病人ID, mlng主页ID) = False Then
            MsgBox "医嘱执行计价数据修正失败，不能继续！", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    If mbyt操作 = Fun_历史转出查询 Then
        cmdOK.Visible = False: cmdOK.Enabled = False
        cmdCancel.Caption = "退出(&E)"
        Me.Caption = "转病区费用变动查询"
    Else
        Me.Caption = "转病区费用变动"
    End If
    
    If ShowPatientInfo(mlng病人ID, mlng主页ID) = False Then Unload Me: Exit Sub
    
    If InitPanel() = False Then Unload Me: Exit Sub
    If InitGrid() = False Then Unload Me: Exit Sub
    If mbyt操作 <> Fun_撤消转病区 Then
        Set objPane = dkpMain.FindPane(Pane_卫材取消)
        If Not objPane Is Nothing Then
            If Not objPane.Closed Then objPane.Close
        End If
    ElseIf mbyt操作 = Fun_撤消转病区 Then
        Set objPane = dkpMain.FindPane(Pane_卫材申请)
        If Not objPane Is Nothing Then
            If Not objPane.Closed Then objPane.Close
        End If
    End If
    
    If mbyt操作 = Fun_历史转出查询 Then
        If LoadHistory(mlng病人ID, mlng主页ID) = False Then Unload Me: Exit Sub
        lblShow(Lbl_变动时间).Caption = lblShow(Lbl_变动时间).Tag
    Else
        lblShow(Lbl_变动时间).Caption = lblShow(Lbl_变动时间).Tag & Format(gobjDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        If mbyt操作 = Fun_撤消转病区 Then
            If LoadFeeData(mlng病人ID, mlng主页ID, mlng原病区id, mlng目标病区id, True, mlng变动id) = False Then Unload Me: Exit Sub
        Else
            If LoadFeeData(mlng病人ID, mlng主页ID, mlng原病区id, mlng目标病区id) = False Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Function LoadHistory(ByVal lng病人ID As Long, ByVal lng主页Id As Long) As Boolean
    '显示历史转病区记录
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    strSql = "Select Distinct 变动时间, 记录状态 From 费用变动记录 Where 病人id = [1] And 主页id = [2] Order By 变动时间"
    Set rsData = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页Id)
    If rsData.EOF Then
        MsgBox "当前病人没有转病区费用变动记录！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    cboDate.Clear
    cboDate.Tag = ""
    Do While Not rsData.EOF
        cboDate.AddItem Format(rsData!变动时间, "yyyy-mm-dd hh:mm:ss")
        cboDate.ItemData(cboDate.NewIndex) = Val(Nvl(rsData!记录状态)) '1-正常变动记录;2-撤销的变动记录
        rsData.MoveNext
    Loop
    cboDate.ListIndex = 0
    
    LoadHistory = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function InitPanel() As Boolean
    '初始化界面布局
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    dkpMain.DestroyAll
    Set objPane = dkpMain.CreatePane(Pane_病人信息, 100, 35, DockTopOf, Nothing)
    objPane.Handle = picBack(Pane_病人信息).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Height = 38
    objPane.MaxTrackSize.Height = 38

    Set objPane = dkpMain.CreatePane(Pane_变动费用, 100, 135, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_变动费用).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    
    Set objPane = dkpMain.CreatePane(Pane_卫材申请, 100, 100, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_卫材申请).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    
    Set objPane = dkpMain.CreatePane(Pane_卫材取消, 100, 80, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_卫材取消).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    
    Set objPane = dkpMain.CreatePane(Pane_功能按钮, 100, 135, DockBottomOf, objPane)
    objPane.Handle = picBack(Pane_功能按钮).hWnd
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Height = 60
    objPane.MaxTrackSize.Height = 60

    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    
    InitPanel = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function InitGrid() As Boolean
    '初始化表格
    Dim strHead(1 To 3) As String, strHeadSub As String
    Dim varData As Variant, varDataSub As Variant
    Dim i As Integer, k As Integer
    
    On Error GoTo ErrHandler
    strHead(1) = "单据性质,1,1050|NO,4,1200|收费细目ID,0,0|收费项目,1,2800|转出,1,1600|转出,7,1000|转入,1,1600|转入,7,1000|单价,7,1300|金额,7,1300"
    strHeadSub = "单据性质,NO,收费细目ID,收费项目,病区,数量,病区,数量,单价,金额"
    strHead(2) = "单据性质,1,1050|NO,4,1200|收费细目ID,0,0|收费项目,1,2800|病区,1,1600|数量,7,1000|操作方式,1,2400"
    strHead(3) = "申请时间,4,2400|申请人,1,1000|NO,4,1200|收费细目ID,0,0|收费项目,1,2800|数量,7,1000|申请部门,1,1600|审核部门,1,1600|申请类别,1,1600"
    
    For k = 1 To 3
        varData = Split(strHead(k), "|")
        If (k = 1) Then varDataSub = Split(strHeadSub, ",")
        
        With vsfGrid(k)
            .Cols = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, i) = Split(varData(i), ",")(0)
                If (k = 1) Then .TextMatrix(1, i) = varDataSub(i)
                .ColAlignment(i) = Split(varData(i), ",")(1)
                .ColWidth(i) = Split(varData(i), ",")(2)
                If k = 1 And (Split(varData(i), ",")(0) = "转出" Or Split(varData(i), ",")(0) = "转入") Then
                    .ColKey(i) = Split(varData(i), ",")(0) & "-" & varDataSub(i)
                Else
                    .ColKey(i) = Split(varData(i), ",")(0)
                End If
            Next
            
            .FixedAlignment(-1) = flexAlignCenterCenter '标题栏文本居中
            If k = 1 Then
                .MergeCells = flexMergeFixedOnly
                .MergeRow(0) = True: .MergeCol(-1) = True
            End If
            
            .ColHidden(.ColIndex("收费细目ID")) = True
        End With
        Call gobjComlib.RestoreFlexState(vsfGrid(k), App.ProductName & "\" & Me.Name & "_" & k)
    Next
    
    InitGrid = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ShowPatientInfo(ByVal lng病人ID As Long, ByVal lng主页Id As Long) As Boolean
    '加载数据
    '入参：
    '
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    strSql = "Select Nvl(a.姓名, b.姓名) As 姓名, Nvl(a.性别, b.性别) As 性别, Nvl(a.年龄, b.年龄) As 年龄," & vbNewLine & _
            "        Nvl(a.住院号, b.住院号) As 住院号, a.主页ID, Nvl(a.费别, b.费别) As 费别," & vbNewLine & _
            "        Nvl(审核标志, 0) As  审核标志, Nvl(a.状态, 0) As 状态" & vbNewLine & _
            " From 病案主页 A, 病人信息 B" & vbNewLine & _
            " Where a.病人id = b.病人id And a.病人id = [1] And a.主页id = [2]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页Id)
    If rsTemp.EOF Then
        MsgBox "未发现病人信息！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lblShow(Lbl_姓名).Caption = lblShow(Lbl_姓名).Tag & Nvl(rsTemp!姓名)
    lblShow(Lbl_性别).Caption = lblShow(Lbl_性别).Tag & Nvl(rsTemp!性别)
    lblShow(Lbl_年龄).Caption = lblShow(Lbl_年龄).Tag & Nvl(rsTemp!年龄)
    lblShow(Lbl_住院号).Caption = lblShow(Lbl_住院号).Tag & Nvl(rsTemp!住院号)
    lblShow(Lbl_住院号).Tag = Val(Nvl(rsTemp!状态)) '住院状态
    lblShow(Lbl_住院次数).Caption = lblShow(Lbl_住院次数).Tag & "第" & Nvl(rsTemp!主页ID) & "次"
    lblShow(Lbl_住院次数).Tag = Val(Nvl(rsTemp!审核标志)) '审核标志
    lblShow(Lbl_费别).Caption = lblShow(Lbl_费别).Tag & Nvl(rsTemp!费别)
    
    ShowPatientInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function LoadFeeData(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    ByVal lng原病区id As Long, ByVal lng目标病区id As Long, _
    Optional ByVal blnCancel As Boolean, Optional ByVal lng目标变动id As Long) As Boolean
    '显示数据
    '入参：
    '   blnCancel - 是否撤销转院
    Dim strSql As String, rsBill As ADODB.Recordset
    Dim lngRow As Long, blnNotData As Boolean
    Dim strDeptNameSql As String, strFeeSql As String
    Dim dbl转出数量 As Double, strAdviceSql As String
    
    Err = 0: On Error GoTo ErrHandler
    blnNotData = True
    '清空界面数据
    vsfGrid(Pane_变动费用).Clear 1
    vsfGrid(Pane_变动费用).Rows = vsfGrid(Pane_变动费用).FixedRows
    vsfGrid(Pane_卫材申请).Clear 1
    vsfGrid(Pane_卫材申请).Rows = vsfGrid(Pane_卫材申请).FixedRows
    vsfGrid(Pane_卫材取消).Clear 1
    vsfGrid(Pane_卫材取消).Rows = vsfGrid(Pane_卫材取消).FixedRows
    
    '费用明细
    strFeeSql = _
        "Select a.No, a.序号, Max(a.收费细目id) As 收费细目id, Max(a.医嘱序号) As 医嘱id, Sum(数量) As 剩余数量," & vbNewLine & _
        "        Max(Decode(a.记录状态, 2, 0, a.费用id)) As 费用id, Max(a.记录状态) As 记录状态, Max(a.执行状态) As 执行状态, Max(a.标准单价) As 标准单价" & vbNewLine & _
        "From (Select a.No, 记录状态, a.医嘱序号, Nvl(a.价格父号, 序号) As 序号, 收费细目id, a.执行状态, Avg(Nvl(a.付数, 1) * a.数次) As 数量," & vbNewLine & _
        "               Sum(a.标准单价) As 标准单价, Max(Decode(a.价格父号, Null, a.Id, 0)) As 费用id" & vbNewLine & _
        "        From 住院费用记录 A, 病人医嘱发送 B, 材料特性 C" & vbNewLine & _
        "        Where a.记录性质 = b.记录性质 And a.医嘱序号 = b.医嘱id And a.No = b.No And a.收费细目id = c.材料id(+) And a.记录性质 = 2 And" & vbNewLine & _
        "              a.执行部门id = [3] And a.医嘱序号 Is Not Null And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
        "              Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0 And Nvl(c.跟踪在用, 0) = 0" & vbNewLine & _
        "        Group By a.No, 记录状态, a.医嘱序号, Nvl(a.价格父号, 序号), 收费细目id, a.执行状态) A" & vbNewLine & _
        "Group By a.No, a.序号" & vbNewLine
    '医嘱信息
    strAdviceSql = _
        "Select b.No, b.医嘱id, b.发送号, c.收费细目id, Nvl(c.费用id, 0) As 费用id, Sum(Decode(Nvl(c.执行状态, 0), 1, c.数量, 0)) As 已执行数" & vbNewLine & _
        "From 病人医嘱记录 A, 病人医嘱发送 B, 医嘱执行计价 C, 材料特性 D" & vbNewLine & _
        "Where a.Id = b.医嘱id And b.发送号 = c.发送号(+) And b.医嘱id = c.医嘱id(+) And b.记录性质 = 2 And Nvl(b.执行状态, 0) <> 1 And" & vbNewLine & _
        "      c.收费细目id = d.材料id(+) And Nvl(d.跟踪在用, 0) = 0 And a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
        "Group By b.No, b.医嘱id, b.发送号, c.收费细目id, Nvl(c.费用id, 0)" & vbNewLine
    
    '病区名称
    strDeptNameSql = _
        " Select Max(Decode(ID, [3], 名称, Null)) As 原病区, Max(Decode(ID, [4], 名称, Null)) As 目标病区" & vbNewLine & _
        " From 部门表 Where ID In ([3], [4])"

    strSql = _
        " Select Decode(a.记录状态, 0, '记帐划价', '记帐') As 单据性质, a.No, Sum(Nvl(剩余数量, 0) - Nvl(已执行数, 0)) As 准退数," & vbNewLine & _
        "        a.收费细目ID, c.名称 As 收费项目, a.标准单价, n.原病区, n.目标病区, Nvl(Sum(剩余数量), 0) As 剩余数量" & vbNewLine & _
        " From (" & strFeeSql & ") A, (" & strAdviceSql & ") B, 收费项目目录 C, (" & strDeptNameSql & ") N" & vbNewLine & _
        " Where a.No = b.No And a.医嘱id = b.医嘱id And a.收费细目id = b.收费细目id And a.收费细目Id = c.Id" & vbNewLine & _
        "       And (a.费用id = b.费用id Or Nvl(b.费用id, 0) = 0) And Nvl(剩余数量, 0) - Nvl(已执行数, 0) > 0" & vbNewLine & _
        " Group By Decode(a.记录状态, 0, '记帐划价', '记帐'), a.No, a.收费细目id, c.名称, a.标准单价, n.原病区, n.目标病区" & vbNewLine & _
        " Having Sum(Nvl(剩余数量, 0) - Nvl(已执行数, 0)) > 0" & vbNewLine & _
        " Order By NO"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "获取费用信息", lng病人ID, lng主页Id, lng原病区id, lng目标病区id)
    If rsBill.RecordCount > 0 Then
        blnNotData = False
        With vsfGrid(Pane_变动费用)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                If Nvl(rsBill!单据性质) = "记帐划价" Then '记账划价单是修改原纪录，因此变动数量是费用的剩余数量
                    dbl转出数量 = Val(Nvl(rsBill!剩余数量))
                Else
                    dbl转出数量 = Val(Nvl(rsBill!准退数))
                End If
                If dbl转出数量 > 0 Then
                    .TextMatrix(lngRow, .ColIndex("单据性质")) = Nvl(rsBill!单据性质)
                    .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                    .TextMatrix(lngRow, .ColIndex("收费细目ID")) = Nvl(rsBill!收费细目ID)
                    .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(rsBill!收费项目)
                    .TextMatrix(lngRow, .ColIndex("转出-病区")) = Nvl(rsBill!原病区)
                    .TextMatrix(lngRow, .ColIndex("转出-数量")) = FormatEx(dbl转出数量, 5)
                    .TextMatrix(lngRow, .ColIndex("转入-病区")) = Nvl(rsBill!目标病区)
                    .TextMatrix(lngRow, .ColIndex("转入-数量")) = FormatEx(dbl转出数量, 5)
                    .TextMatrix(lngRow, .ColIndex("单价")) = Format(Val(Nvl(rsBill!标准单价)), gSysPara.Price_Decimal.strFormt_VB)
                    .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsBill!标准单价)) * dbl转出数量, gSysPara.Money_Decimal.strFormt_VB)
                    
                    lngRow = lngRow + 1
                End If
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
        
    '卫材销账申请
    '跟踪在用的卫材执行了的也允许销账申请
    strSql = _
        "With 住院费用 As (" & _
        " Select a.No, a.序号, Max(a.收费细目id) As 收费细目id, Sum(数量) As 剩余数量, Max(Decode(a.记录状态, 2, 0, a.费用id)) As 费用id," & vbNewLine & _
        "         Max(a.记录状态) As 记录状态" & vbNewLine & _
        " From (Select a.No, 记录状态, Nvl(a.价格父号, 序号) As 序号, 收费细目id, Avg(Nvl(a.付数, 1) * a.数次) As 数量," & vbNewLine & _
        "                Max(Decode(a.价格父号, Null, a.Id, 0)) As 费用id" & vbNewLine & _
        "         From 住院费用记录 A, 病人医嘱发送 B, 材料特性 C" & vbNewLine & _
        "         Where a.记录性质 = b.记录性质 And a.医嘱序号 = b.医嘱id And a.No = b.No And a.记录性质 = 2 And a.执行部门id = [3] And" & vbNewLine & _
        "               a.医嘱序号 Is Not Null And a.病人id = [1] And a.主页id = [2] And Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0 And" & vbNewLine & _
        "               a.收费细目id = c.材料id And Nvl(c.跟踪在用, 0) = 1" & vbNewLine & _
        "         Group By a.No, 记录状态, a.医嘱序号, Nvl(a.价格父号, 序号), 收费细目id, a.执行状态) A" & vbNewLine & _
        " Group By a.No, a.序号)"
    
    strSql = _
        " Select Decode(a.记录状态, 0, '记帐划价', '记帐') As 单据性质, a.No, a.收费细目id, 准退数, b.名称 As 收费项目, n.原病区" & vbNewLine & _
        " From (" & vbNewLine & _
        "        " & strSql & vbNewLine & _
        "        Select Nvl(Sum(a.剩余数量), 0) - Nvl(Sum(b.数量), 0) As 准退数, a.No, Max(a.记录状态) As 记录状态, a.收费细目id" & vbNewLine & _
        "        From 住院费用 A," & vbNewLine & _
        "             (Select b.费用id, Nvl(Sum(b.数量), 0) As 数量" & vbNewLine & _
        "               From 住院费用 A, 病人费用销帐 B" & vbNewLine & _
        "               Where a.费用id = B.费用id And Nvl(B.状态, 0) = 0" & vbNewLine & _
        "               Group By b.费用id" & vbNewLine & _
        "               Having Nvl(Sum(b.数量), 0) <> 0) B" & vbNewLine & _
        "        Where a.费用id = b.费用id(+)" & vbNewLine & _
        "        Group By a.No, a.收费细目id" & vbNewLine & _
        "        Having Nvl(Sum(a.剩余数量), 0) - Nvl(Sum(b.数量), 0) <> 0) A, 收费项目目录 B,(" & strDeptNameSql & ") N" & vbNewLine & _
        " Where a.收费细目ID = B.ID" & vbNewLine & _
        " Order By NO, 收费细目id"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "获取费用信息", lng病人ID, lng主页Id, lng原病区id, lng目标病区id)
    If rsBill.RecordCount > 0 Then
        blnNotData = False
        With vsfGrid(Pane_卫材申请)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                .TextMatrix(lngRow, .ColIndex("单据性质")) = Nvl(rsBill!单据性质)
                .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                .TextMatrix(lngRow, .ColIndex("收费细目ID")) = Nvl(rsBill!收费细目ID)
                .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(rsBill!收费项目)
                .TextMatrix(lngRow, .ColIndex("病区")) = Nvl(rsBill!原病区)
                .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(Nvl(rsBill!准退数), 5)
                If Nvl(rsBill!单据性质) = "记帐划价" Then
                    .TextMatrix(lngRow, .ColIndex("操作方式")) = "禁止转病区"
                    '红色字体标记
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                Else
                    .TextMatrix(lngRow, .ColIndex("操作方式")) = "销账申请"
                End If
                
                lngRow = lngRow + 1
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
    
    '销账申请取消
    If blnCancel Then
        strSql = "Select ID From 病人变动记录 Where 病人ID = [1] And 主页ID = [2] And 开始原因 = 15 And 终止时间 Is Null"
        Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "获取原变动ID", lng病人ID, lng主页Id)
        If rsBill.EOF Then
            MsgBox "未找到病人的原始变动记录，禁止操作！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSql = _
            " Select a.申请时间, a.申请人, b.No, a.收费细目ID, c.名称 As 收费项目, Sum(a.数量) As 数量, e.名称 As 申请部门, f.名称 As 审核部门," & vbNewLine & _
            "        Decode(a.申请类别, 1, '已执行', '未执行') As 申请类别" & vbNewLine & _
            " From 病人费用销帐 A, 费用变动记录 B, 收费项目目录 C, 部门表 E, 部门表 F" & vbNewLine & _
            " Where a.费用id = b.费用id And a.收费细目id = c.Id And a.申请部门id = e.Id And a.审核部门id = f.Id" & vbNewLine & _
            "       And b.原变动id = [1] And b.目标变动ID = [2] And b.状态 = 2 And a.状态 In (0, 2)" & vbNewLine & _
            " Group By a.申请时间, a.申请人, b.No, a.收费细目id, c.名称, e.名称, f.名称, Decode(a.申请类别, 1, '已执行', '未执行')" & vbNewLine & _
            " Order By No, 收费细目ID"
        Set rsBill = gobjDatabase.OpenSQLRecord(strSql, "获取销账申请取消数据", lng目标变动id, Val(Nvl(rsBill!ID)))
        If rsBill.RecordCount > 0 Then
            blnNotData = False
            With vsfGrid(Pane_卫材取消)
                .Redraw = flexRDNone
                .Rows = .FixedRows + rsBill.RecordCount
                lngRow = .FixedRows
                Do While Not rsBill.EOF
                    .TextMatrix(lngRow, .ColIndex("申请时间")) = Format(Nvl(rsBill!申请时间), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngRow, .ColIndex("申请人")) = Nvl(rsBill!申请人)
                    .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                    .TextMatrix(lngRow, .ColIndex("收费细目ID")) = Nvl(rsBill!收费细目ID)
                    .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(rsBill!收费项目)
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(Val(Nvl(rsBill!数量)), 5)
                    .TextMatrix(lngRow, .ColIndex("申请部门")) = Nvl(rsBill!申请部门)
                    .TextMatrix(lngRow, .ColIndex("审核部门")) = Nvl(rsBill!审核部门)
                    .TextMatrix(lngRow, .ColIndex("申请类别")) = Nvl(rsBill!申请类别)
                    
                    lngRow = lngRow + 1
                    rsBill.MoveNext
                Loop
                .Redraw = flexRDBuffered
            End With
        End If
    End If
    
    If blnNotData Then
        'MsgBox "当前病人没有可转出的费用！", vbInformation, gstrSysName
        mblnOK = True
        Exit Function
    End If
    LoadFeeData = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim k As Integer
    
    For k = 1 To 3
        Call gobjComlib.SaveFlexState(vsfGrid(k), App.ProductName & "\" & Me.Name & "_" & k)
    Next
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case Pane_病人信息
        shp病人信息.Move 10, 10, picBack(Index).ScaleWidth - 30, picBack(Index).ScaleHeight - 10
        lblShow(Lbl_年龄).Left = lblShow(Lbl_性别).Left + lblShow(Lbl_性别).Width + 500
        lblShow(Lbl_住院号).Left = lblShow(Lbl_年龄).Left + lblShow(Lbl_年龄).Width + 500
        lblShow(Lbl_住院次数).Left = lblShow(Lbl_住院号).Left + lblShow(Lbl_住院号).Width + 500
        lblShow(Lbl_费别).Left = lblShow(Lbl_住院次数).Left + lblShow(Lbl_住院次数).Width + 500
    Case Pane_变动费用, Pane_卫材申请, Pane_卫材取消
        lblShow(Index).Left = 50
        lblShow(Index).Top = 50
        With vsfGrid(Index)
            .Left = 10
            .Top = lblShow(Index).Top + lblShow(Index).Height + 30
            .Width = picBack(Index).ScaleWidth - .Left - 20
            .Height = picBack(Index).ScaleHeight - .Top
        End With
        
        If Index = Pane_变动费用 Then
            picCboBack.Left = lblShow(Index).Left + lblShow(Index).Width
            picCboBack.Top = lblShow(Index).Top - 50
        End If
    Case Pane_功能按钮
        cmdCancel.Left = picBack(Index).ScaleWidth - cmdCancel.Width - 800
        cmdCancel.Top = (picBack(Index).ScaleHeight - cmdCancel.Height) / 2 - 50
        cmdOK.Top = cmdCancel.Top
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    End Select
End Sub

Public Function Upgrade医嘱执行计价执行状态(ByVal lng病人ID As Long, ByVal lng主页Id As Long) As Boolean
    '功能：修正"医嘱执行计价.执行状态"
    '入参：
    '   lng病人ID
    '   lng主页ID
    '返回：已修正则返回True，否则返回False
    '问题号:99715
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Distinct a.Id As 医嘱id, b.No" & vbNewLine & _
        " From 病人医嘱记录 A, 病人医嘱发送 B, 医嘱执行计价 C" & vbNewLine & _
        " Where a.Id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号" & vbNewLine & _
        "       And b.记录性质 = 2 And a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
        "       And c.执行状态 Is Null"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "判断医嘱执行计价执行状态是否已修正", lng病人ID, lng主页Id)
    If rsTemp.RecordCount = 0 Then
        Upgrade医嘱执行计价执行状态 = True
        Exit Function
    End If
    
    '修正数据
    Do While Not rsTemp.EOF
        'Zl_医嘱执行计价_修正(
        strSql = "Zl_医嘱执行计价_修正("
        '  医嘱id_In   病人医嘱执行.医嘱id%Type,
        strSql = strSql & "" & Nvl(rsTemp!医嘱ID) & ","
        '  No_In       病人医嘱发送.No%Type,
        strSql = strSql & "'" & Nvl(rsTemp!NO) & "',"
        '  记录性质_In 病人医嘱发送.记录性质%Type
        strSql = strSql & "" & 2 & ")"
        gobjDatabase.ExecuteProcedure strSql, "修正数据"
        
        rsTemp.MoveNext
    Loop
    
    Upgrade医嘱执行计价执行状态 = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub cboDate_Click()
    Dim strSql As String, rsBill As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandler
    If cboDate.Tag = cboDate.Text Then Exit Sub
    cboDate.Tag = cboDate.Text
    If cboDate.ListIndex < 0 Then Exit Sub
    
    '清空界面数据
    vsfGrid(Pane_变动费用).Clear 1
    vsfGrid(Pane_变动费用).Rows = vsfGrid(Pane_变动费用).FixedRows
    vsfGrid(Pane_卫材申请).Clear 1
    vsfGrid(Pane_卫材申请).Rows = vsfGrid(Pane_卫材申请).FixedRows
    
    strSql = _
        " Select Decode(h.记录状态, 0, '记帐划价', '记帐') As 单据性质, a.No, a.收费细目ID, b.名称 As 收费项目," & vbNewLine & _
        "        c.名称 As 原病区, d.名称 As 目标病区, a.数量, a.单价, a.实收金额" & vbNewLine & _
        " From 住院费用记录 H, 费用变动记录 A, 收费项目目录 B, 部门表 C, 部门表 D" & vbNewLine & _
        " Where h.Id = a.费用id And a.收费细目id = b.Id And a.原病区id = c.Id And a.目标病区id = d.Id" & vbNewLine & _
        "       And a.病人id = [1] And a.主页id = [2] And a.变动时间 = [3] And 状态 In (0, 1)" & vbNewLine & _
        " Order By No"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, CDate(cboDate.Text))
    If rsBill.RecordCount > 0 Then
        With vsfGrid(Pane_变动费用)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                .TextMatrix(lngRow, .ColIndex("单据性质")) = Nvl(rsBill!单据性质)
                .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                .TextMatrix(lngRow, .ColIndex("收费细目ID")) = Nvl(rsBill!收费细目ID)
                .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(rsBill!收费项目)
                .TextMatrix(lngRow, .ColIndex("转出-病区")) = Nvl(rsBill!原病区)
                .TextMatrix(lngRow, .ColIndex("转出-数量")) = FormatEx(Val(Nvl(rsBill!数量)), 5)
                .TextMatrix(lngRow, .ColIndex("转入-病区")) = Nvl(rsBill!目标病区)
                .TextMatrix(lngRow, .ColIndex("转入-数量")) = FormatEx(Val(Nvl(rsBill!数量)), 5)
                .TextMatrix(lngRow, .ColIndex("单价")) = Format(Nvl(rsBill!单价), gSysPara.Price_Decimal.strFormt_VB)
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Nvl(rsBill!实收金额), gSysPara.Money_Decimal.strFormt_VB)
                
                lngRow = lngRow + 1
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
    
    strSql = _
        " Select Decode(h.记录状态, 0, '记帐划价', '记帐') As 单据性质, a.No, a.收费细目ID, b.名称 As 收费项目, c.名称 As 原病区," & vbNewLine & _
        "        Sum(a.数量) As 数量, Decode(a.状态, 2, '销帐申请', '取消申请') As 操作方式" & vbNewLine & _
        " From 住院费用记录 H, 费用变动记录 A, 收费项目目录 B, 部门表 C" & vbNewLine & _
        " Where h.Id = a.费用id And a.收费细目id = b.Id And a.原病区id = c.Id And a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
        "       And a.变动时间 = [3] And a.状态 In (2, 3)" & vbNewLine & _
        " Group By Decode(h.记录状态, 0, '记帐划价', '记帐'), a.No, a.收费细目id, b.名称, c.名称, a.数量, Decode(a.状态, 2, '销帐申请', '取消申请')" & vbNewLine & _
        " Order By No"
    Set rsBill = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng主页ID, CDate(cboDate.Text))
    If rsBill.RecordCount > 0 Then
        With vsfGrid(Pane_卫材申请)
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsBill.RecordCount
            lngRow = .FixedRows
            Do While Not rsBill.EOF
                .TextMatrix(lngRow, .ColIndex("单据性质")) = Nvl(rsBill!单据性质)
                .TextMatrix(lngRow, .ColIndex("NO")) = Nvl(rsBill!NO)
                .TextMatrix(lngRow, .ColIndex("收费细目ID")) = Nvl(rsBill!收费细目ID)
                .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(rsBill!收费项目)
                .TextMatrix(lngRow, .ColIndex("病区")) = Nvl(rsBill!原病区)
                .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(Val(Nvl(rsBill!数量)), 5)
                .TextMatrix(lngRow, .ColIndex("操作方式")) = Nvl(rsBill!操作方式)
                
                lngRow = lngRow + 1
                rsBill.MoveNext
            Loop
            .Redraw = flexRDBuffered
        End With
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim lngRow As Long
    Dim byt住院状态 As Byte, byt审核标志 As Byte
    
    On Error GoTo ErrHandler
    
     '卫材销账申请存在未审核的不允许转病区
    With vsfGrid(Pane_卫材申请)
        For lngRow = .FixedRows To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("操作方式")) = "禁止转病区" Then
                MsgBox "单据“" & .TextMatrix(lngRow, .ColIndex("NO")) & "”还未进行审核，禁止转病区操作！", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
    End With
    
    byt住院状态 = Val(lblShow(Lbl_住院号).Tag)
    If gSysPara.bln未入科禁止记账 And byt住院状态 = 1 Then
        MsgBox "病人未入科，禁止对病人相关费用的操作，因此禁止转病区操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    If gSysPara.byt病人审核方式 = 1 Then
        byt审核标志 = Val(lblShow(Lbl_住院次数).Tag)
        If byt审核标志 = 1 Then
            MsgBox "该病人目前正在审核费用，不能进行费用相关调整，因此禁止转病区操作！", vbInformation, gstrSysName
            Exit Sub
        ElseIf byt审核标志 = 2 Then
            MsgBox "该病人目前已经完成了费用审核，不能进行费用相关调整，因此禁止转病区操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mbyt操作 = Fun_转病区申请 Then '转病区申请退出
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    
    Set mcllSQL = New Collection
    '  --功能:病人转病区费用的转入，转出处理
    '  --入参:操作_In: 0-病区变动,1-撤消病区变动
    '  --   1、操作_IN=0(病区变动)时
    '  --       变动ID_In: 原病区的变动记录的ID
    '  --       原病区ID_IN：原病区ID
    '  --       目标病区ID_IN:目标病区ID
    '  --   2、操作_IN=1(撤消病区变动)时
    '  --       变动ID_In: 恢复的原始病区的变动记录的ID
    '  --       原病区ID_IN：被撤消的病区ID
    '  --       目标病区ID_IN:恢复的原始病区ID
    '  --转入，转出规则:
    '  --1.病区执行的非药品和卫生材料，处理规则为
    '  --   1)将原记录进行销帐处理
    '  --   2)新增一条新病区的费用，病人科室，发生时间不变
    '  --2.病区执行的药品和卫生材料
    '  --   这个卫材退的处理在转病区时的界面中进行确认(可以打印核查清单)，在转病区发起的时候确认。
    '  --   a)卫材在原病区通过销帐申请来处理，新病区手工计卫材；
    '  --   b)撤消转病区时，自动撤消销帐申请，如果已经销帐审核了，则询问提示并且不作卫材费用处理，手工去处理。
    'Zl_Turntoward_Fee
    strSql = "Zl_Turntoward_Fee("
    '(
    '  操作_In       Number,
    strSql = strSql & "" & mbyt操作 & ","
    '  病人id_In     病案主页.病人id%Type,
    strSql = strSql & "" & mlng病人ID & ","
    '  主页id_In     病案主页.病人id%Type,
    strSql = strSql & "" & mlng主页ID & ","
    '  变动id_In   病人变动记录.Id%Type,
    strSql = strSql & "" & mlng变动id & ","
    '  原病区id_In   病案主页.当前病区id%Type,
    strSql = strSql & "" & mlng原病区id & ","
    '  目标病区id_In 病案主页.当前病区id%Type,
    strSql = strSql & "" & mlng目标病区id & ","
    '  操作员编号_In 住院费用记录.操作员编号%Type,
    strSql = strSql & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "')"
    '  变动时间_In   住院费用记录.登记时间%Type := Null
    
    mcllSQL.Add strSql
    mblnOK = True
    Unload Me
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

