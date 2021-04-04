VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmAuditLogManage 
   BackColor       =   &H80000005&
   Caption         =   "重要操作变动日志管理"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   ControlBox      =   0   'False
   Icon            =   "frmAuditLogManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmAuditLogManage.frx":6852
   ScaleHeight     =   7305
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList img16 
      Left            =   9315
      Top             =   2580
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
            Picture         =   "frmAuditLogManage.frx":6D4B
            Key             =   "system"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditLogManage.frx":D5AD
            Key             =   "program"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgModuleType 
      Left            =   120
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   51
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditLogManage.frx":13E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditLogManage.frx":1735C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAuditLogConfig 
      BackColor       =   &H80000005&
      FillColor       =   &H80000000&
      Height          =   6060
      Left            =   1110
      ScaleHeight     =   6000
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   5955
      Width           =   8235
      Begin MSComctlLib.TreeView tvwAuditLogConfig 
         Height          =   3030
         Left            =   90
         TabIndex        =   32
         Top             =   495
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   5345
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "停用(&T)"
         Height          =   350
         Left            =   4410
         TabIndex        =   24
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "启用(&S)"
         Height          =   350
         Left            =   3315
         TabIndex        =   23
         Top             =   45
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   585
         TabIndex        =   21
         Top             =   75
         Width           =   2000
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditLogConfig 
         Height          =   1680
         Left            =   2760
         TabIndex        =   18
         Top             =   495
         Width           =   4770
         _cx             =   8414
         _cy             =   2963
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAuditLogManage.frx":1A852
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
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "查找"
         Height          =   180
         Left            =   105
         TabIndex        =   22
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox picAuditLogList 
      BackColor       =   &H80000005&
      Height          =   5955
      Left            =   915
      ScaleHeight     =   5895
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   225
      Width           =   7815
      Begin VB.Frame fraDescription 
         BackColor       =   &H80000005&
         Caption         =   "操作说明"
         Height          =   2300
         Left            =   4425
         TabIndex        =   20
         Top             =   3630
         Width           =   3150
         Begin VB.TextBox txtInstructions 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1920
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   255
            Width           =   2835
         End
      End
      Begin VB.Frame fraNote 
         BackColor       =   &H80000005&
         Caption         =   "操作内容"
         Height          =   2300
         Left            =   345
         TabIndex        =   19
         Top             =   2790
         Width           =   3150
         Begin VB.TextBox txtOperationContent 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1920
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   255
            Width           =   2835
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditLogList 
         Height          =   2730
         Left            =   -10
         TabIndex        =   17
         Top             =   -10
         Width           =   3360
         _cx             =   5927
         _cy             =   4815
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAuditLogManage.frx":1A950
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
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3540
         Left            =   4125
         ScaleHeight     =   3540
         ScaleWidth      =   3495
         TabIndex        =   11
         Top             =   0
         Width           =   3500
         Begin VB.ListBox lisShowList 
            Appearance      =   0  'Flat
            Height          =   1470
            Left            =   915
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   420
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.ComboBox cboFunction 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1659
            Width           =   2385
         End
         Begin VB.ComboBox cboSystem 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   897
            Width           =   2385
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "查找(&F)"
            Height          =   350
            Left            =   2175
            TabIndex        =   10
            Top             =   3135
            Width           =   1100
         End
         Begin VB.ComboBox cboWorkStation 
            Height          =   300
            Left            =   915
            TabIndex        =   5
            Top             =   135
            Width           =   2385
         End
         Begin VB.ComboBox cboModule 
            Height          =   300
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1278
            Width           =   2385
         End
         Begin VB.ComboBox cboUserName 
            Height          =   300
            Left            =   915
            TabIndex        =   6
            Top             =   516
            Width           =   2385
         End
         Begin MSComCtl2.DTPicker dtpDateEnd 
            Height          =   315
            Left            =   915
            TabIndex        =   9
            Top             =   2700
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   105381891
            CurrentDate     =   37029
         End
         Begin MSComCtl2.DTPicker dtpDateStart 
            Height          =   315
            Left            =   915
            TabIndex        =   8
            Top             =   2040
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   105381891
            CurrentDate     =   37029
         End
         Begin VB.Label lblFunction 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "操作功能"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   31
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label lblSystem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "所属系统"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   29
            Top             =   945
            Width           =   720
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "至"
            Height          =   180
            Left            =   915
            TabIndex        =   16
            Top             =   2430
            Width           =   180
         End
         Begin VB.Label lblWorkStation 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "客户端"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   330
            TabIndex        =   15
            Top             =   210
            Width           =   540
         End
         Begin VB.Label lblModule 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "操作模块"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   14
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label lblUserName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "用户名"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   330
            TabIndex        =   13
            Top             =   570
            Width           =   540
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "操作时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   165
            TabIndex        =   12
            Top             =   2085
            Width           =   720
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   1395
      Left            =   60
      TabIndex        =   2
      Top             =   615
      Width           =   1560
      _Version        =   589884
      _ExtentX        =   2752
      _ExtentY        =   2461
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdLogClear 
      Caption         =   "日志清理(&C)"
      Height          =   375
      Left            =   8430
      TabIndex        =   1
      Top             =   150
      Width           =   1290
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作日志管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   1440
   End
End
Attribute VB_Name = "frmAuditLogManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsAuditLog As ADODB.Recordset '记录按条件查询出来的日志数据，主要用于展示详细操作内容和操作说明
Private mrsModuleList As ADODB.Recordset '记录模块及其启停信息，主要用于模块的查找
Private mrsWorkStation As ADODB.Recordset '记录客户端信息，主要用于客户端模糊查找
Private mrsUserName As ADODB.Recordset '记录用户名信息，主要用于用户名的模糊查找
Private mrsSysProgFun As ADODB.Recordset '记录系统、模块及功能的对应关系
Private mlngCurPos As Long '当前查找树形结构的位置

Private Enum AuditLogList
    VLL_用户名 = 0
    VLL_人员 = 1
    VLL_部门 = 2
    VLL_工作站 = 3
    VLL_操作类型 = 4
    VLL_系统编号 = 5
    VLL_所属系统 = 6
    VLL_操作模块编号 = 7
    VLL_操作模块 = 8
    VLL_操作功能 = 9
    VLL_操作时间 = 10
    VLL_操作内容 = 11
    VLL_操作说明 = 12
    VLL_详细操作内容 = 13
    VLL_详细操作说明 = 14
End Enum

Private Enum AuditLogConfig
    VLC_ID = 0
    VLC_所属系统 = 1
    VLC_模块名称 = 2
    VLC_功能名称 = 3
    VLC_说明 = 4
    VLC_需审核 = 5
    VLC_状态 = 6
    VLC_状态标记 = 7
End Enum

Private Sub cboFunction_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '点击回车，光标移到下一个控件
        dtpDateStart.SetFocus
    End If
End Sub

Private Sub cboModule_Click()
    If cboModule.Text = "" Then
        '若模块选择为空，则在操作功能下拉框中展示对应系统中的全部功能数据
        If cboSystem.Text = "" Then
            mrsSysProgFun.Filter = ""
        Else
            mrsSysProgFun.Filter = "系统 = " & Split(cboSystem.Text, "-")(0)
        End If
    Else
        '若模块选择不为空，则在操作功能下拉框中仅展示该模块中的全部功能数据
        If cboSystem.Text = "" Then
            mrsSysProgFun.Filter = "模块 = " & Split(cboModule.Text, "-")(0)
        Else
            mrsSysProgFun.Filter = "系统 = " & Split(cboSystem.Text, "-")(0) & " And 模块 = " & Split(cboModule.Text, "-")(0)
        End If
    End If
    
    '填充操作功能数据
    mrsSysProgFun.Sort = "Id"
    cboFunction.Clear
    cboFunction.addItem ""
    Do While Not mrsSysProgFun.EOF
        cboFunction.addItem mrsSysProgFun!id & "-" & mrsSysProgFun!功能
        mrsSysProgFun.MoveNext
    Loop
End Sub

Private Sub cboModule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '点击回车，光标移到下一个控件
        cboFunction.SetFocus
    End If
End Sub

Private Sub cboSystem_Click()
    Dim strLastModuleNo As String

    If cboSystem.Text = "" Then
        '若系统选择为空，则在操作模块及操作功能下拉框中展示全部的数据
        mrsSysProgFun.Filter = ""
    Else
        '若系统选择为应用系统，那么在操作模块及操作功能下拉框中仅展示对应系统相关数据
        mrsSysProgFun.Filter = "系统 = " & Split(cboSystem.Text, "-")(0)
    End If
    
    '填充操作模块数据
    mrsSysProgFun.Sort = "系统, 模块"
    cboModule.Clear
    cboModule.addItem ""
    Do While Not mrsSysProgFun.EOF
        If mrsSysProgFun!系统 & "-" & mrsSysProgFun!模块 <> strLastModuleNo Then
            cboModule.addItem mrsSysProgFun!模块 & "-" & mrsSysProgFun!模块名称
            strLastModuleNo = mrsSysProgFun!系统 & "-" & mrsSysProgFun!模块
        End If
        mrsSysProgFun.MoveNext
    Loop
    
    '填充操作功能数据
    mrsSysProgFun.Sort = "Id"
    cboFunction.Clear
    cboFunction.addItem ""
    Do While Not mrsSysProgFun.EOF
        cboFunction.addItem mrsSysProgFun!id & "-" & mrsSysProgFun!功能
        mrsSysProgFun.MoveNext
    Loop
End Sub

Private Sub cboSystem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '点击回车，光标移到下一个控件
        cboModule.SetFocus
    End If
End Sub

Private Sub cboUserName_Change()
    '模糊查找，并将结果显示到listBox中
    '查找模式：按名称查找
    If cboUserName.Locked Then Exit Sub
    If cboUserName.Text <> "" Then
        lisShowList.Top = cboUserName.Top + cboUserName.Height
        lisShowList.Visible = True
    Else
        lisShowList.Visible = False
        Exit Sub
    End If
    mrsUserName.Filter = "用户名 like '%" & cboUserName.Text & "%'"
    lisShowList.Clear
    With mrsUserName
        lisShowList.Height = 210 * .RecordCount
        If lisShowList.Height > 1470 Then lisShowList.Height = 1470
        Do While Not .EOF
            lisShowList.addItem !用户名
            .MoveNext
        Loop
    End With
End Sub

Private Sub cboUserName_DropDown()
    lisShowList.Visible = False
End Sub

Private Sub cboUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    '按下“下方向键”,将焦点转移到列表中
    If KeyCode = 40 And lisShowList.Visible And lisShowList.ListCount <> 0 Then
        cboUserName.Locked = True
        lisShowList.SetFocus
        lisShowList.ListIndex = 0
    ElseIf KeyCode = 13 Then
        '当点击回车，光标移到下一个控件
        cboSystem.SetFocus
    Else
        KeyCode = 0
    End If
End Sub

Private Sub cboUserName_LostFocus()
    If lisShowList.ListCount = 0 Then lisShowList.Visible = False
End Sub

Private Sub cboWorkStation_Change()
    '模糊查找，并将结果显示到listBox中
    '查找模式：按名称查找
    If cboWorkStation.Locked Then Exit Sub
    If cboWorkStation.Text <> "" Then
        lisShowList.Top = cboWorkStation.Top + cboWorkStation.Height
        lisShowList.Visible = True
    Else
        lisShowList.Visible = False
        Exit Sub
    End If
    mrsWorkStation.Filter = "工作站 like '%" & cboWorkStation.Text & "%' or 简码 like '%" & cboWorkStation.Text & "%'"
    lisShowList.Clear
    With mrsWorkStation
        lisShowList.Height = 210 * .RecordCount
        If lisShowList.Height > 1470 Then lisShowList.Height = 1470
        Do While Not .EOF
            lisShowList.addItem !工作站
            .MoveNext
        Loop
    End With
End Sub

Private Sub cboWorkStation_DropDown()
    lisShowList.Visible = False
End Sub

Private Sub cboWorkStation_KeyDown(KeyCode As Integer, Shift As Integer)
    '按下“下方向键”,将焦点转移到列表中
    If KeyCode = 40 And lisShowList.Visible And lisShowList.ListCount <> 0 Then
        cboWorkStation.Locked = True
        lisShowList.SetFocus
        lisShowList.ListIndex = 0
    ElseIf KeyCode = 13 Then
        '当点击回车，光标移到下一个控件
        cboUserName.SetFocus
    Else
        KeyCode = 0
    End If
End Sub

Private Sub cboWorkStation_LostFocus()
    If lisShowList.ListCount = 0 Then lisShowList.Visible = False
End Sub

Private Sub cmdFind_Click()
    Call FillAuditLog
End Sub

'填充日志数据
Private Sub FillAuditLog()
    Dim strSQL As String
    Dim lngSystemNo As Long, lngFunctionNo As String
    Dim strModuleNo As String
    Dim i As Long
    
    On Error GoTo errH
    If cboWorkStation.Text <> "" Then strSQL = " And A.工作站 = [1]"
    If cboUserName.Text <> "" Then strSQL = strSQL & " And a.用户名 = [2]"
    If cboSystem.Text <> "" Then
        lngSystemNo = Split(cboSystem.Text, "-")(0)
        strSQL = strSQL & " And Nvl(f.系统,0) = [3]"
    End If
    If cboModule.Text <> "" Then
        strModuleNo = Split(cboModule.Text, "-")(0)
        strSQL = strSQL & " And f.模块 = [4]"
    End If
    If cboFunction.Text <> "" Then
        lngFunctionNo = Split(cboFunction.Text, "-")(0)
        strSQL = strSQL & " And f.Id = [5]"
    End If
    strSQL = strSQL & " And a.操作时间 between [6] and [7]"
    strSQL = "Select a.用户名, d.姓名, e.名称 部门, a.工作站, a.操作时间, Decode(a.操作类型, 1, '增加', 2, '修改', '删除') 操作类型, 0 系统, '服务器管理工具' 所属系统, f.模块," & vbNewLine & _
            "       g.标题 操作模块, f.功能, a.操作内容, a.操作说明" & vbNewLine & _
            "From Zlauditlog A, 上机人员表 B, 部门人员 C, 人员表 D, 部门表 E, Zlauditlogconfig F, zlSvrTools G" & vbNewLine & _
            "Where a.用户名 = b.用户名 And b.人员id = d.Id And b.人员id = c.人员id And c.部门id = e.Id And c.缺省 = 1 And a.日志Id = f.Id And" & vbNewLine & _
            "      f.模块 = g.编号 And f.系统 Is Null" & strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.用户名, d.姓名, e.名称 部门, a.工作站, a.操作时间, Decode(a.操作类型, 1, '增加', 2, '修改', '删除') 操作类型, f.系统, h.名称 所属系统, f.模块," & vbNewLine & _
            "       g.标题 操作模块, f.功能, a.操作内容, a.操作说明" & vbNewLine & _
            "From Zlauditlog A, 上机人员表 B, 部门人员 C, 人员表 D, 部门表 E, Zlauditlogconfig F, zlPrograms G, zlSystems H" & vbNewLine & _
            "Where a.用户名 = b.用户名 And b.人员id = d.Id And b.人员id = c.人员id And c.部门id = e.Id And c.缺省 = 1 And a.日志Id = f.Id And" & vbNewLine & _
            "      f.模块 = g.序号 And f.系统 = g.系统 And f.系统 = h.编号" & strSQL
    frmMDIMain.stbThis.Panels(2).Text = "正在查找！"
    Set mrsAuditLog = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, cboWorkStation.Text, cboUserName.Text, lngSystemNo, strModuleNo, lngFunctionNo, _
                    CDate(Format(dtpDateStart.value, "YYYY-MM-DD") & " 00:00:00"), CDate(Format(dtpDateEnd.value, "YYYY-MM-DD") & " 23:59:59"))
    With mrsAuditLog
        .Sort = "操作时间"
        vsfAuditLogList.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfAuditLogList.TextMatrix(i, VLL_用户名) = !用户名
            vsfAuditLogList.TextMatrix(i, VLL_人员) = !姓名
            vsfAuditLogList.TextMatrix(i, VLL_部门) = !部门
            vsfAuditLogList.TextMatrix(i, VLL_工作站) = !工作站
            vsfAuditLogList.TextMatrix(i, VLL_操作类型) = !操作类型
            vsfAuditLogList.TextMatrix(i, VLL_系统编号) = !系统
            vsfAuditLogList.TextMatrix(i, VLL_所属系统) = !所属系统
            vsfAuditLogList.TextMatrix(i, VLL_操作模块编号) = !模块
            vsfAuditLogList.TextMatrix(i, VLL_操作模块) = !操作模块
            vsfAuditLogList.TextMatrix(i, VLL_操作功能) = !功能
            vsfAuditLogList.TextMatrix(i, VLL_操作时间) = !操作时间
            vsfAuditLogList.TextMatrix(i, VLL_操作内容) = IIf(Len(!操作内容) > 50, Mid(!操作内容, 1, 50) & "...", !操作内容)
            vsfAuditLogList.TextMatrix(i, VLL_操作说明) = IIf(Len(!操作说明) > 50, Mid(!操作说明 & "", 1, 50) & "...", !操作说明 & "")
            vsfAuditLogList.TextMatrix(i, VLL_详细操作内容) = !操作内容
            vsfAuditLogList.TextMatrix(i, VLL_详细操作说明) = !操作说明 & ""
            .MoveNext
        Next
        frmMDIMain.stbThis.Panels(2).Text = "共查找到“" & .RecordCount & "”条数据！"
        vsfAuditLogList.Tag = frmMDIMain.stbThis.Panels(2).Text
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdLogClear_Click()
    If frmAuditLogClear.ShowMe() Then
        '判断当前界面是否有日志数据，如果有，执行刷新操作，否则不刷新
        If vsfAuditLogList.Rows > 1 Then
            Call FillAuditLog
        End If
    End If
End Sub

Private Sub cmdStart_Click()
    On Error GoTo errH
    '开启该模块的日志
    Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",1)", "启用模块日志")
    vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_状态标记) = 1
    Call RecUpdate(mrsModuleList, "Id = " & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID), "是否启用", 1)
    vsfAuditLogConfig.Cell(flexcpPicture, vsfAuditLogConfig.Row, VLC_状态) = imgModuleType.ListImages(2).Picture
    cmdStart.Enabled = False
    CmdStop.Enabled = True
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub CmdStop_Click()
    On Error GoTo errH
    '停用该模块的日志
    Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",0)", "停用模块日志")
    vsfAuditLogConfig.Cell(flexcpPicture, vsfAuditLogConfig.Row, VLC_状态) = imgModuleType.ListImages(1).Picture
    vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_状态标记) = 0
    Call RecUpdate(mrsModuleList, "Id = " & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID), "是否启用", 0)
    cmdStart.Enabled = True
    CmdStop.Enabled = False
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub dtpDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '当点击回车，光标移到下一个控件
        cmdFind.SetFocus
    End If
End Sub

Private Sub dtpDateStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        '当点击回车，光标移到下一个控件
        dtpDateEnd.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    '对tebControl控件进行初始化
    Call InitTabControl
    
    '填充基础数据，主要为日志查找部分下拉框数据
    Call FillBaseData
    
    '填充模块启停数据
    Call FillModuleTree
End Sub

'==============================================================================
'=功能： 初始Tab控件
'==============================================================================
Private Function InitTabControl() As Boolean
    Dim objTabItem As TabControlItem
On Error GoTo errH
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .OneNoteColors = True
            .DisableLunaColors = True
        End With
        '第一页
        Set objTabItem = .InsertItem(0, "日志查看", picAuditLogList.hwnd, 0)
        '第二页
        .InsertItem 1, "日志启停", picAuditLogConfig.hwnd, 0
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With

    InitTabControl = True

    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

'填充基础数据
Private Sub FillBaseData()
    Dim rsTmp As ADODB.Recordset
    Dim lngLastSystemNo As Long
    Dim strLastModuleNo As String

    On Error GoTo errH
    '填充工作站数据
    gstrSQL = "Select 工作站,zlspellcode(工作站) 简码 From zlClients"
    Set mrsWorkStation = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
    Do While Not mrsWorkStation.EOF
        cboWorkStation.addItem mrsWorkStation!工作站
        mrsWorkStation.MoveNext
    Loop
    
    '填充用户名数据
    gstrSQL = "Select 用户名 From 上机人员表"
    Set mrsUserName = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
    Do While Not mrsUserName.EOF
        cboUserName.addItem mrsUserName!用户名
        mrsUserName.MoveNext
    Loop
    
    '记录系统、模块及功能的对应关系，方便查找功能的使用
    gstrSQL = "Select a.Id, 0 系统, '服务器管理工具' 系统名称, a.模块, b.标题 模块名称, a.功能" & vbNewLine & _
            "From Zlauditlogconfig A, zlSvrTools B" & vbNewLine & _
            "Where a.模块 = b.编号 And a.系统 Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.Id, a.系统, b.名称 系统名称, a.模块, c.标题 模块名称, a.功能" & vbNewLine & _
            "From Zlauditlogconfig A, zlSystems B, zlPrograms C" & vbNewLine & _
            "Where a.系统 = b.编号 And a.模块 = c.序号 And b.编号 = c.系统"
    Set mrsSysProgFun = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption)
    
    '填充应用系统和管理工具数据
    mrsSysProgFun.Sort = "系统"
    lngLastSystemNo = -1
    cboSystem.addItem ""
    Do While Not mrsSysProgFun.EOF
        If mrsSysProgFun!系统 <> lngLastSystemNo Then
            cboSystem.addItem mrsSysProgFun!系统 & "-" & mrsSysProgFun!系统名称
            lngLastSystemNo = mrsSysProgFun!系统
        End If
        mrsSysProgFun.MoveNext
    Loop
    
    '填充操作模块数据
    mrsSysProgFun.Sort = "系统, 模块"
    cboModule.addItem ""
    Do While Not mrsSysProgFun.EOF
        If mrsSysProgFun!系统 & "-" & mrsSysProgFun!模块 <> strLastModuleNo Then
            cboModule.addItem mrsSysProgFun!模块 & "-" & mrsSysProgFun!模块名称
            strLastModuleNo = mrsSysProgFun!系统 & "-" & mrsSysProgFun!模块
        End If
        mrsSysProgFun.MoveNext
    Loop
    
    '填充操作功能数据
    mrsSysProgFun.Sort = "Id"
    cboFunction.addItem ""
    Do While Not mrsSysProgFun.EOF
        cboFunction.addItem mrsSysProgFun!id & "-" & mrsSysProgFun!功能
        mrsSysProgFun.MoveNext
    Loop
    
    '填充时间数据
    dtpDateStart.value = CurrentDate()
    dtpDateEnd.value = dtpDateStart.value
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'填充树形结构系统及模块信息
Private Sub FillModuleTree()
    Dim lngLastSystemNo As Long   '最后一次添加的系统编号
    Dim lngLaseProgNo As Long   '最后一次添加的模块的模块编号
    Dim objNode As Node
    Dim i As Long

    On Error GoTo errH
    gstrSQL = "Select *" & vbNewLine & _
                "From (Select a.Id, 0 系统, '服务器管理工具' 系统名称, a.模块, b.标题 模块名称, zlSpellCode(b.标题) 简码, a.功能, a.说明, a.是否需审核, a.是否启用" & vbNewLine & _
                "       From Zlauditlogconfig A, zlSvrTools B" & vbNewLine & _
                "       Where a.模块 = b.编号 And a.系统 Is Null" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select a.Id, a.系统, c.名称 系统名称, a.模块, b.标题 模块名称, zlSpellCode(b.标题) 简码, a.功能, a.说明, a.是否需审核, a.是否启用" & vbNewLine & _
                "       From Zlauditlogconfig A, zlPrograms B, zlSystems C" & vbNewLine & _
                "       Where a.系统 = b.系统 And a.模块 = b.序号 And a.系统 = c.编号)" & vbNewLine & _
                "Order By 系统, 模块"
    '如果树形结构中有数据，说明已经进行了初始化了，即本次调用为查找功能调用
    If tvwAuditLogConfig.Nodes.Count = 0 Then
        Set mrsModuleList = CopyNewRec(gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption))
        lngLastSystemNo = -1
        '填充数据
        With mrsModuleList
            Do While Not .EOF
                If !系统 <> lngLastSystemNo Then
                    Set objNode = tvwAuditLogConfig.Nodes.Add(, , "K_" & !系统, !系统名称, "system")
                    objNode.Expanded = True
                    lngLastSystemNo = !系统
                    Set objNode = tvwAuditLogConfig.Nodes.Add("K_" & lngLastSystemNo, tvwChild, "K_" & lngLastSystemNo & "_" & !模块, !模块名称, "program")
                    objNode.Tag = !简码
                    lngLaseProgNo = !模块
                Else
                    If !模块 <> lngLaseProgNo Then
                        Set objNode = tvwAuditLogConfig.Nodes.Add("K_" & lngLastSystemNo, tvwChild, "K_" & lngLastSystemNo & "_" & !模块, !模块名称, "program")
                        objNode.Tag = !简码
                        lngLaseProgNo = !模块
                    End If
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then
                tvwAuditLogConfig.Nodes(1).Child.Selected = True
                tvwAuditLogConfig.Tag = tvwAuditLogConfig.SelectedItem.Key
                Call tvwAuditLogConfig_NodeClick(tvwAuditLogConfig.SelectedItem)
            End If
        End With
    Else
        '将光标定位在要查找的数据行上
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).BackColor = &H80000005
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).ForeColor = &H80000012
        If mlngCurPos > tvwAuditLogConfig.Nodes.Count Then mlngCurPos = 1
        For i = mlngCurPos To tvwAuditLogConfig.Nodes.Count
            Set objNode = tvwAuditLogConfig.Nodes(i)
            If objNode.Tag <> "" Then
                If objNode.Text Like "*" & txtFind.Text & "*" Or objNode.Tag Like "*" & UCase(txtFind.Text) & "*" Then
                    objNode.Expanded = True
                    objNode.Selected = True
                    objNode.BackColor = &H8000000D
                    objNode.ForeColor = &H80000005
                    tvwAuditLogConfig.Tag = tvwAuditLogConfig.SelectedItem.Key
                    mlngCurPos = i
                    Call tvwAuditLogConfig_NodeClick(objNode)
                    Exit For
                ElseIf i = tvwAuditLogConfig.Nodes.Count Then
                    mlngCurPos = 0
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'填充模块启停信息
Private Sub FillModuleList()
    Dim i As Long

    On Error GoTo errH
    '将查询到的信息填充到界面上
    With mrsModuleList
        vsfAuditLogConfig.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfAuditLogConfig.TextMatrix(i, VLC_ID) = !id
            vsfAuditLogConfig.TextMatrix(i, VLC_所属系统) = !系统名称
            vsfAuditLogConfig.TextMatrix(i, VLC_模块名称) = !模块名称
            vsfAuditLogConfig.TextMatrix(i, VLC_功能名称) = !功能
            vsfAuditLogConfig.TextMatrix(i, VLC_说明) = !说明 & ""
            If !是否需审核 = 1 Then
                vsfAuditLogConfig.TextMatrix(i, VLC_需审核) = "√"
            ElseIf !是否需审核 = 2 Then
                vsfAuditLogConfig.TextMatrix(i, VLC_需审核) = "×"
            Else
                vsfAuditLogConfig.TextMatrix(i, VLC_需审核) = ""
            End If
            vsfAuditLogConfig.TextMatrix(i, VLC_状态标记) = !是否启用
            vsfAuditLogConfig.Cell(flexcpPicture, i, VLC_状态) = imgModuleType.ListImages(!是否启用 + 1).Picture
            .MoveNext
        Next
        If .RecordCount > 0 Then
            vsfAuditLogConfig.Row = 1
            Call vsfAuditLogConfig_Click
        End If
    End With
    If VScrollVisible(vsfAuditLogConfig) Then
        vsfAuditLogConfig.ColWidth(VLC_说明) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                            - vsfAuditLogConfig.ColWidth(VLC_状态) - vsfAuditLogConfig.ColWidth(VLC_功能名称) - 350
    Else
        vsfAuditLogConfig.ColWidth(VLC_说明) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                            - vsfAuditLogConfig.ColWidth(VLC_状态) - vsfAuditLogConfig.ColWidth(VLC_功能名称) - 100
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Private Sub Form_Resize()
    On Error Resume Next
    tbcPage.Left = 0
    tbcPage.Width = Me.ScaleWidth
    tbcPage.Top = 520
    tbcPage.Height = Me.ScaleHeight - tbcPage.Top
    cmdLogClear.Top = 80
    cmdLogClear.Left = Me.ScaleWidth - 200 - cmdLogClear.Width
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsAuditLog = Nothing
    Set mrsModuleList = Nothing
    mlngCurPos = 0
End Sub

Private Sub lisShowList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        '当按下“上方向键”或“下方向键”时，同步更新cboWorkStation中的值
        If cboWorkStation.Locked Then
            cboWorkStation.Text = lisShowList.List(lisShowList.ListIndex)
        Else
            cboUserName.Text = lisShowList.List(lisShowList.ListIndex)
        End If
    ElseIf KeyCode = 13 Then
        '当点击回车，选定当前选中的数据
        If cboWorkStation.Locked Then
            lisShowList.Visible = False
            cboWorkStation.Locked = False
            cboWorkStation.SetFocus
        Else
            lisShowList.Visible = False
            cboUserName.Locked = False
            cboUserName.SetFocus
        End If
    End If
End Sub

Private Sub picAuditLogConfig_Resize()
    On Error Resume Next
    vsfAuditLogConfig.Width = picAuditLogConfig.Width - vsfAuditLogConfig.Left - 150
    vsfAuditLogConfig.Height = picAuditLogConfig.Height - vsfAuditLogConfig.Top - 10
    tvwAuditLogConfig.Height = vsfAuditLogConfig.Height
    CmdStop.Left = picAuditLogConfig.Width - CmdStop.Width - 180
    cmdStart.Left = CmdStop.Left - cmdStart.Width
    If VScrollVisible(vsfAuditLogConfig) Then
        vsfAuditLogConfig.ColWidth(VLC_说明) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                            - vsfAuditLogConfig.ColWidth(VLC_状态) - vsfAuditLogConfig.ColWidth(VLC_功能名称) - 350
    Else
        vsfAuditLogConfig.ColWidth(VLC_说明) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                            - vsfAuditLogConfig.ColWidth(VLC_状态) - vsfAuditLogConfig.ColWidth(VLC_功能名称) - 100
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub picAuditLogList_Resize()
    On Error Resume Next
    vsfAuditLogList.Width = picAuditLogList.Width - picFind.Width
    vsfAuditLogList.Height = picAuditLogList.Height
    picFind.Left = vsfAuditLogList.Width
    fraDescription.Top = picAuditLogList.Height - fraDescription.Height - 200
    fraDescription.Left = picFind.Left + 150
    fraNote.Top = fraDescription.Top - fraNote.Height - 200
    fraNote.Left = fraDescription.Left
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "日志查看" Then
        cmdLogClear.Visible = True
        frmMDIMain.stbThis.Panels(2).Text = vsfAuditLogList.Tag
    Else
        cmdLogClear.Visible = False
        frmMDIMain.stbThis.Panels(2).Text = ""
    End If
End Sub

Private Sub tvwAuditLogConfig_NodeClick(ByVal Node As MSComctlLib.Node)
    '填充右侧列表数据
    If tvwAuditLogConfig.Tag <> "" Then
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).BackColor = &H80000005
        tvwAuditLogConfig.Nodes(tvwAuditLogConfig.Tag).ForeColor = &H80000012
    End If
    Node.BackColor = &H8000000D
    Node.ForeColor = &H80000005
    tvwAuditLogConfig.Tag = Node.Key
    If tvwAuditLogConfig.SelectedItem.Parent Is Nothing Then
        mrsModuleList.Filter = "系统 = " & Split(tvwAuditLogConfig.SelectedItem.Key, "_")(1)
        vsfAuditLogConfig.MergeCol(VLC_模块名称) = True
    Else
        mrsModuleList.Filter = "系统 = " & Split(tvwAuditLogConfig.SelectedItem.Key, "_")(1) & " And 模块 = '" & Split(tvwAuditLogConfig.SelectedItem.Key, "_")(2) & "'"
        vsfAuditLogConfig.MergeCol(VLC_模块名称) = False
    End If
    Call FillModuleList
End Sub

Private Sub txtFind_Change()
    mlngCurPos = 0
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String

    If KeyCode = vbKeyReturn Then
        mlngCurPos = mlngCurPos + 1
        Call FillModuleTree
    End If
End Sub

Private Sub vsfAuditLogConfig_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If VScrollVisible(vsfAuditLogConfig) Then
        vsfAuditLogConfig.ColWidth(VLC_说明) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                                - vsfAuditLogConfig.ColWidth(VLC_功能名称) - vsfAuditLogConfig.ColWidth(VLC_状态) - 350
        vsfAuditLogConfig.ColWidth(VLC_模块名称) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_说明) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                                - vsfAuditLogConfig.ColWidth(VLC_功能名称) - vsfAuditLogConfig.ColWidth(VLC_状态) - 350
        vsfAuditLogConfig.ColWidth(VLC_功能名称) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                                - vsfAuditLogConfig.ColWidth(VLC_说明) - vsfAuditLogConfig.ColWidth(VLC_状态) - 350
    Else
        vsfAuditLogConfig.ColWidth(VLC_说明) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                                - vsfAuditLogConfig.ColWidth(VLC_功能名称) - vsfAuditLogConfig.ColWidth(VLC_状态) - 100
        vsfAuditLogConfig.ColWidth(VLC_模块名称) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_说明) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                                - vsfAuditLogConfig.ColWidth(VLC_功能名称) - vsfAuditLogConfig.ColWidth(VLC_状态) - 100
        vsfAuditLogConfig.ColWidth(VLC_功能名称) = vsfAuditLogConfig.Width - vsfAuditLogConfig.ColWidth(VLC_模块名称) - vsfAuditLogConfig.ColWidth(VLC_需审核) _
                                                - vsfAuditLogConfig.ColWidth(VLC_说明) - vsfAuditLogConfig.ColWidth(VLC_状态) - 100
    End If
End Sub

Private Sub vsfAuditLogConfig_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = VLC_说明 Or Col = VLC_需审核 Then
        Cancel = True
    End If
End Sub

Private Sub vsfAuditLogConfig_Click()
    If vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.RowSel, VLC_状态标记) = 1 Then
        cmdStart.Enabled = False
        CmdStop.Enabled = True
    Else
        cmdStart.Enabled = True
        CmdStop.Enabled = False
    End If
End Sub

Private Sub vsfAuditLogConfig_DblClick()
    On Error GoTo errH
    With vsfAuditLogConfig
        If .MouseRow <> .Row Then Exit Sub
        '只有当鼠标双击“状态”一列时，才进行启停操作，以免产生误操作
        If .ColSel = VLC_状态 Then
            If .TextMatrix(.RowSel, VLC_状态标记) = 1 Then
                Call CmdStop_Click
            Else
                Call cmdStart_Click
            End If
        ElseIf .ColSel = VLC_需审核 Then
            If .TextMatrix(.RowSel, VLC_需审核) = "√" Then
                .TextMatrix(.RowSel, VLC_需审核) = "×"
                Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",Null,2)", "设置为无需审核")
            ElseIf .TextMatrix(.RowSel, VLC_需审核) = "×" Then
                .TextMatrix(.RowSel, VLC_需审核) = "√"
                Call ExecuteProcedure("Zl_Zlauditlogconfig_Update(" & vsfAuditLogConfig.TextMatrix(vsfAuditLogConfig.Row, VLC_ID) & ",Null,1)", "设置为需审核")
            End If
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub vsfAuditLogConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sinLeft As Single, sinRight As Single
    Dim strTip As String
    
    With vsfAuditLogConfig
        sinLeft = .ColWidth(VLC_模块名称) + .ColWidth(VLC_功能名称) + .ColWidth(VLC_说明)
        sinRight = sinLeft + .ColWidth(VLC_需审核)
        If X >= sinLeft And X <= sinRight And Y <= 260 Then
            strTip = "只有一些特别重要的操作才需进行审核操作。" & vbNewLine & _
                           "若开启此参数，当用户使用对应功能时，若用户为普通管理员，则需要进行管理员身份验证并填写操作说明，若用户为系统所有者，则只需填写操作说明。" & vbNewLine & _
                           "若不开启此参数，将无需进行身份验证及填写操作说明。"
            Call ShowTipInfo(.hwnd, strTip, True)
        Else
            Call ShowTipInfo(.hwnd, "")
        End If
    End With
End Sub

Private Sub vsfAuditLogList_Click()
    '当点击每一行时，在右下方显示详细的操作内容和操作说明
    With vsfAuditLogList
        If .MouseRow <> .Row Or .Row < 1 Then Exit Sub
        '因为界面上记录的是被截取过后的数据，故要想显示详细的操作内容和操作说明，将直接展示已经加载的隐藏数据
        txtOperationContent.Text = .TextMatrix(.Row, VLL_详细操作内容)
        txtInstructions.Text = .TextMatrix(.Row, VLL_详细操作说明)
    End With
End Sub
