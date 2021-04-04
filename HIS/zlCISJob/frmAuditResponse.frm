VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAuditResponse 
   AutoRedraw      =   -1  'True
   Caption         =   "病案审查反馈"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "frmAuditResponse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11385
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   5280
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "病人"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "住院号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "床号"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "年龄"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "部门"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "拼音简码"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox picPati 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2270
         Picture         =   "frmAuditResponse.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "选择病人(F4)"
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox txtPati 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "筛选病人"
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
         Left            =   150
         TabIndex        =   7
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   2730
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   1170
   End
   Begin RichTextLib.RichTextBox txtResponse 
      Height          =   765
      Left            =   720
      TabIndex        =   1
      Top             =   6225
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1349
      _Version        =   393217
      BackColor       =   14737632
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAuditResponse.frx":0680
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
   Begin MSComctlLib.ImageList img16 
      Left            =   1740
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":071D
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":0CB7
            Key             =   "处理状态_等待处理"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":1251
            Key             =   "处理状态_处理暂存"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":7AB3
            Key             =   "处理状态_等待复查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":804D
            Key             =   "处理状态_结束"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":85E7
            Key             =   "对象_医嘱"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":8B81
            Key             =   "对象_病历"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":911B
            Key             =   "对象_护理"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":96B5
            Key             =   "对象_首页"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":9C4F
            Key             =   "对象_报告"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":A1E9
            Key             =   "对象_文件"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":A783
            Key             =   "对象_路径"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7185
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAuditResponse.frx":10FE5
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17171
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
   Begin RichTextLib.RichTextBox txtNote 
      Height          =   765
      Left            =   5670
      TabIndex        =   2
      Top             =   6255
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      MaxLength       =   255
      Appearance      =   0
      TextRTF         =   $"frmAuditResponse.frx":11877
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
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   75
      ScaleHeight     =   5325
      ScaleWidth      =   11190
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   570
      Width           =   11190
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   5145
         Left            =   60
         TabIndex        =   0
         Top             =   90
         Width           =   11010
         _Version        =   589884
         _ExtentX        =   19420
         _ExtentY        =   9075
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   165
      Top             =   135
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAuditResponse.frx":11914
      Left            =   1245
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   630
      Top             =   135
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmAuditResponse.frx":11928
   End
End
Attribute VB_Name = "frmAuditResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Closed(ByVal DataChange As Boolean)
Public Event OpenObject(ByVal PatiID As Long, ByVal PageID As Long, ByVal ObjectType As Integer, ByVal ObjectID As String)

Private Enum ICON_ID
    conIcon_UnCheck = 1
    conIcon_Check = 2
    conIcon_UnSelect = 3
    conIcon_Select = 4
End Enum
Private Enum MENU_ID
    conMenu_FilterLable = 0
    conMenu_Submit = 1
    conMenu_Random = 2
    conMenu_Await = 3
    conMenu_Done = 4
    conMenu_DateLable = 5
    conMenu_DateInput = 6
    conMenu_Pati = 7
    
    conMenu_Refresh = 90
    conMenu_OpenData = 91
    
    conMenu_Save = 92
    conMenu_Commit = 93
    conMenu_CommitOne = 94
    conMenu_Cancel = 95
    
    conMenu_Help = 98
    conMenu_Exit = 99
    
    conMenu_Col = 100
    
    conMenu_AllCollapse = 201
    conMenu_AllExpend = 202
    conMenu_CurCollapse = 203
    conMenu_CurExpend = 204
End Enum

Private Enum COLUMN_PATI
    pcol_住院号 = 1
    pcol_床号 = 2
    pcol_性别 = 3
    pcol_年龄 = 4
    pcol_部门 = 5
    pcol_拼音简码 = 6
End Enum

Private Enum COLUMN_ID
    col_状态 = 0 '显示状态图标
    col_姓名 = 1
    col_住院号 = 2
    col_床号 = 3
    col_性别 = 4
    col_年龄 = 5
    col_部门 = 6
    
    col_反馈对象 = 7 '显示对象图标
    col_反馈意见 = 8
    col_补充说明 = 9
    col_处理期限 = 10
    col_反馈人 = 11
    col_反馈时间 = 12
    
    col_处理说明 = 13
    col_处理人 = 14
    col_处理时间 = 15
    col_分值 = 16
    
    col_病人Id = 17
    col_主页ID = 18
    col_反馈ID = 19
    col_相关ID = 20
    col_对象ID = 21
    col_子文档ID = 22 '新版病历
End Enum

Private mstrPrivs As String
Private mlngDeptID As Long '病区/科室ID
Private mintDeptType As Integer '0-按科室显示，1-按病区显示
Private mintDataType As Integer '0-医生站用,1-护士站用
Private mblnICU As Boolean '是否非本科的ICU科室

Private Type FilterCond
    提交审查 As Boolean
    随机抽查 As Boolean
    未处理 As Boolean
    已处理 As Boolean
    开始时间 As Date
    结束时间 As Date
End Type
Private mvarCond As FilterCond
Private mblnEditing As Boolean
Private mintPreTime As Integer
Private mblnOpen As Boolean
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal lngDeptID As Long, ByVal intDeptType As Integer, _
    ByVal blnICU As Boolean, ByVal intDataType As Integer, ByVal strPrivs As String) As Boolean
    mlngDeptID = lngDeptID
    mintDeptType = intDeptType
    mblnICU = blnICU
    mintDataType = intDataType
    mstrPrivs = strPrivs
        
    If mblnOpen Then
        '重新刷新数据
        '###
            
        If Me.WindowState = vbMinimized Then
            Me.WindowState = vbNormal
        End If
    End If
    
    Me.Show , frmParent
End Function

Private Sub cboTime_Click()
    Dim curDate As Date
    
    If cboTime.ListIndex = mintPreTime Then Exit Sub
    
    curDate = zlDatabase.Currentdate
    
    Select Case cboTime.Text
    Case "今天"
        mvarCond.开始时间 = Format(curDate, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "昨天"
        mvarCond.开始时间 = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
    Case "最近三天"
        mvarCond.开始时间 = Format(curDate - 2, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近一周"
        mvarCond.开始时间 = Format(curDate - 7, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近两周"
        mvarCond.开始时间 = Format(curDate - 14, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近一月"
        mvarCond.开始时间 = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "[指定..]"
        If Not frmSelectTime.ShowMe(Me, mvarCond.开始时间, mvarCond.结束时间, cboTime) Then
            '取消时恢复原来的选择
            Call Cbo.SetIndex(cboTime.hwnd, mintPreTime)
            rptData.SetFocus: Exit Sub
        Else
            rptData.SetFocus
        End If
    End Select
        
    cboTime.ToolTipText = "范围：" & Format(mvarCond.开始时间, "yyyy-MM-dd") & " 至 " & Format(mvarCond.结束时间, "yyyy-MM-dd")
    mintPreTime = cboTime.ListIndex
    Me.Refresh
    
    '刷新数据
    Call RefreshData
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    Select Case Control.ID
        Case conMenu_Submit
            Control.IconId = IIf(mvarCond.提交审查, conIcon_Check, conIcon_UnCheck)
            Control.Checked = IIf(mvarCond.提交审查, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_Random
            Control.IconId = IIf(mvarCond.随机抽查, conIcon_Check, conIcon_UnCheck)
            Control.Checked = IIf(mvarCond.随机抽查, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_Await
            Control.IconId = IIf(mvarCond.未处理, conIcon_Select, conIcon_UnSelect)
            Control.Checked = IIf(mvarCond.未处理, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_Done
            Control.IconId = IIf(mvarCond.已处理, conIcon_Select, conIcon_UnSelect)
            Control.Checked = IIf(mvarCond.已处理, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_DateLable, conMenu_DateInput
            Control.Visible = mvarCond.已处理
            Control.Enabled = Not mblnEditing
        Case conMenu_OpenData '打开定位
            If InStr(mstrPrivs, "审查反馈处理") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = False
                If rptData.SelectedRows.Count > 0 Then
                    With rptData.SelectedRows(0)
                        If Not .GroupRow And .Childs.Count = 0 Then
                            blnEnabled = .Record(col_状态).Value = 1 Or .Record(col_状态).Value = 2
                        End If
                    End With
                End If
                Control.Enabled = blnEnabled And Not mblnEditing
            End If
        Case conMenu_Refresh
            Control.Enabled = Not mblnEditing
        Case conMenu_Save '暂存
            Control.Enabled = mblnEditing
            Control.Visible = mvarCond.未处理 '已处理的，不能再执行暂存
        Case conMenu_CommitOne  '完成单条
            If mvarCond.未处理 Then
                Control.Caption = "完成单条"
                Control.ToolTipText = "将当前暂存行的反馈处理提交再次审查"
            Else
                Control.Caption = "保存"
                Control.ToolTipText = "保存修改的处理情况"
            End If
            
            Control.Enabled = mblnEditing
            
            '暂存的，可以直接完成单条
            If mblnEditing = False And mvarCond.未处理 Then
                If rptData.SelectedRows.Count > 0 Then
                    With rptData.SelectedRows(0)
                        If Not .GroupRow And .Childs.Count = 0 Then
                            Control.Enabled = .Record(col_状态).Value = 1 And .Record(col_处理人).Value <> ""
                        End If
                    End With
                End If
            End If
        Case conMenu_Commit     '完成（提交所有暂存）
            Control.Visible = mvarCond.未处理
            Control.Enabled = Not mblnEditing
        
        Case conMenu_Cancel '取消
            Control.Enabled = mblnEditing
        Case conMenu_Col + 1 To conMenu_Col + 99 '显示/隐藏列
            Control.Checked = rptData.Columns.Find(Val(Control.Parameter)).Visible
            Control.Enabled = Not mblnEditing
        Case conMenu_CurExpend '展开当前组
            blnEnabled = False
            If rptData.SelectedRows.Count > 0 Then
                If rptData.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptData.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled And Not mblnEditing
        Case conMenu_CurCollapse '折叠当前组
            blnEnabled = False
            If rptData.SelectedRows.Count > 0 Then
                If rptData.SelectedRows(0).GroupRow Then
                    blnEnabled = rptData.SelectedRows(0).Expanded
                ElseIf Not rptData.SelectedRows(0).ParentRow Is Nothing Then
                    If rptData.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptData.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled And Not mblnEditing
    End Select
    
    If mblnEditing Then
        txtResponse.Enabled = False
        picData.Enabled = False
    Else
        txtResponse.Enabled = True
        picData.Enabled = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objRow As ReportRow
    
    Select Case Control.ID
        Case conMenu_Submit
            If mvarCond.提交审查 And Not mvarCond.随机抽查 Then Exit Sub
            mvarCond.提交审查 = Not mvarCond.提交审查
            Call RefreshData
        Case conMenu_Random
            If mvarCond.随机抽查 And Not mvarCond.提交审查 Then Exit Sub
            mvarCond.随机抽查 = Not mvarCond.随机抽查
            Call RefreshData
        Case conMenu_Await
            If mvarCond.未处理 Then Exit Sub
            mvarCond.未处理 = True: mvarCond.已处理 = False
            Call RefreshData
        Case conMenu_Done
            If mvarCond.已处理 Then Exit Sub
            mvarCond.已处理 = True: mvarCond.未处理 = False
            Call RefreshData
        Case conMenu_Refresh
            Call RefreshData
        Case conMenu_OpenData '打开定位
            Call rptData_RowDblClick(rptData.SelectedRows(0), rptData.SelectedRows(0).Record(col_反馈意见))
        Case conMenu_Save   '暂存
            If SaveData(True) Then
                mblnEditing = False
                cbsMain.RecalcLayout
                rptData.SetFocus
            End If
            
        Case conMenu_CommitOne  '完成单条
            If Trim(txtNote.Text) = "" And mvarCond.未处理 Then
                Call MsgBox("请输入处理说明。", vbInformation, gstrSysName)
                If txtNote.Enabled And txtNote.Visible Then txtNote.SetFocus
                Exit Sub
            End If
            If SaveData(False) Then
                mblnEditing = False
                cbsMain.RecalcLayout
                rptData.SetFocus
            End If
        Case conMenu_Commit '完成所有
            
            Call SaveAllPaseData    '其中包含界面刷新
            
        Case conMenu_Cancel
            If txtNote.Text <> rptData.SelectedRows(0).Record(col_处理说明).Value Then
                If MsgBox("确实要取消编辑吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            txtNote.Text = rptData.SelectedRows(0).Record(col_处理说明).Value
            mblnEditing = False
            cbsMain.RecalcLayout
            rptData.SetFocus
        Case conMenu_Col + 1 To conMenu_Col + 99 '显示/隐藏列
            rptData.Columns.Find(Val(Control.Parameter)).Visible = Not rptData.Columns.Find(Val(Control.Parameter)).Visible
        Case conMenu_CurCollapse '折叠当前组
            If rptData.SelectedRows.Count > 0 Then
                If rptData.SelectedRows(0).GroupRow Then
                    rptData.SelectedRows(0).Expanded = False
                ElseIf Not rptData.SelectedRows(0).ParentRow Is Nothing Then
                    If rptData.SelectedRows(0).ParentRow.GroupRow Then
                        rptData.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '因折叠定位到分组上,不会自动激活该事件
            Call rptData_SelectionChanged
        Case conMenu_CurExpend '展开当前组
            If rptData.SelectedRows.Count > 0 Then
                rptData.SelectedRows(0).Expanded = True
            End If
        Case conMenu_AllCollapse '折叠所有组
            For Each objRow In rptData.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '因折叠定位到分组上,不会自动激活该事件
            Call rptData_SelectionChanged
        Case conMenu_AllExpend '展开所有组
            For Each objRow In rptData.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_Help
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.picData
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = txtResponse.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = txtNote.hwnd
    End If
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = imgMain.Icons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Refresh, "刷新")
            objControl.BeginGroup = True
            objControl.ToolTipText = "刷新当前选择的数据"
        Set objControl = .Add(xtpControlButton, conMenu_OpenData, "打开")
            objControl.ToolTipText = "打开反馈对象"
            
        Set objControl = .Add(xtpControlButton, conMenu_Save, "暂存")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_Commit, "完成")
        objControl.ToolTipText = "将你暂存的反馈处理全部提交再次审查"
        Set objControl = .Add(xtpControlButton, conMenu_CommitOne, "完成单条")
        objControl.ToolTipText = "将当前暂存行的反馈处理提交再次审查"
        objControl.ToolTipText = "保存修改的处理情况"
        
        Set objControl = .Add(xtpControlButton, conMenu_Cancel, "取消")
            
        Set objControl = .Add(xtpControlButton, conMenu_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "退出")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
        
    Set objBar = cbsMain.Add("过滤栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_FilterLable, "过滤条件")
        
        Set objControl = .Add(xtpControlButton, conMenu_Submit, "提交审查")
        Set objControl = .Add(xtpControlButton, conMenu_Random, "随机抽查")
        
        Set objControl = .Add(xtpControlButton, conMenu_Await, "未处理")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Done, "已处理")
        
        Set objControl = .Add(xtpControlLabel, conMenu_DateLable, "处理时间")
            objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, conMenu_DateInput, "处理时间")
            objCustom.Handle = cboTime.hwnd
            
        Set objCustom = .Add(xtpControlCustom, conMenu_Pati, "筛选病人")
            objCustom.Handle = picPati.hwnd
            picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    
    '热键绑定:注意不能和系统的文本编辑热键冲突
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help
        .Add 0, vbKeyF5, conMenu_Refresh
        .Add 0, vbKeyF3, conMenu_OpenData
        .Add 0, vbKeyF2, conMenu_CommitOne
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add 0, vbKeyEscape, conMenu_Cancel
        .Add FALT, vbKeyX, conMenu_Exit
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If txtNote.Enabled And txtNote.Visible And txtNote.Locked = False Then
            txtNote.SetFocus
            Call txtNote_GotFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim curDate As Date
    
    Call InitCommandBar
    
    '缺省医嘱时间
    cboTime.AddItem "今天"
    cboTime.AddItem "昨天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近两周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指定..]"
    mintPreTime = 0
    Call Cbo.SetIndex(cboTime.hwnd, 0)
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 320, 100, DockBottomOf, Nothing)
    objPane.Title = "反馈意见"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set objPane = Me.dkpMain.CreatePane(2, 320, 100, DockRightOf, objPane)
    objPane.Title = "处理说明"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    '其他
    '-----------------------------------------------------
    mblnOpen = True
    mblnOK = False
    
    '缺省条件
    curDate = zlDatabase.Currentdate
    With mvarCond
        .提交审查 = Val(zlDatabase.GetPara("提交审查反馈", glngSys, IIf(mintDataType = 0, p住院医生站, p住院护士站), "1")) <> 0
        .随机抽查 = Val(zlDatabase.GetPara("随机抽查反馈", glngSys, IIf(mintDataType = 0, p住院医生站, p住院护士站), "0")) <> 0
        .未处理 = True: .已处理 = False
        .开始时间 = Format(curDate, "yyyy-MM-dd 00:00:00")
        .结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    End With
    
    '刷新数据
    Call RefreshData
    mblnEditing = False
        
    '------------
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnSetup As Boolean

    If mblnEditing Then
        If MsgBox("确实要取消编辑并退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("提交审查反馈", IIf(mvarCond.提交审查, 1, 0), glngSys, IIf(mintDataType = 0, p住院医生站, p住院护士站), blnSetup)
    Call zlDatabase.SetPara("随机抽查反馈", IIf(mvarCond.随机抽查, 1, 0), glngSys, IIf(mintDataType = 0, p住院医生站, p住院护士站), blnSetup)
    Call SaveWinState(Me, App.ProductName)

    mblnOpen = False
    RaiseEvent Closed(mblnOK)
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptData
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(col_状态, "状态", 75, False)
            objCol.Alignment = xtpAlignmentCenter
            objCol.Icon = img16.ListImages("图标").Index - 1
        Set objCol = .Columns.Add(col_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(col_住院号, "住院号", 62, True)
        Set objCol = .Columns.Add(col_床号, "床号", 40, True)
        Set objCol = .Columns.Add(col_性别, "性别", 30, True)
        Set objCol = .Columns.Add(col_年龄, "年龄", 30, True)
        Set objCol = .Columns.Add(col_部门, IIf(mintDeptType = 0, "科室", "病区"), 70, True)
        
        Set objCol = .Columns.Add(col_反馈对象, "对象", 18, False)
            objCol.Alignment = xtpAlignmentCenter
            objCol.AllowRemove = False
            objCol.Icon = img16.ListImages("图标").Index - 1
        Set objCol = .Columns.Add(col_反馈意见, "反馈意见", 200, True)
            objCol.AllowRemove = False
            objCol.Groupable = False
        Set objCol = .Columns.Add(col_补充说明, "补充说明", 120, True)
        Set objCol = .Columns.Add(col_处理期限, "处理期限", 80, True)
        Set objCol = .Columns.Add(col_反馈人, "反馈人", 50, True)
        Set objCol = .Columns.Add(col_反馈时间, "反馈时间", 80, True)
        
        Set objCol = .Columns.Add(col_处理说明, "处理说明", 200, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(col_处理人, "处理人", 50, True)
        Set objCol = .Columns.Add(col_处理时间, "处理时间", 80, True)
        Set objCol = .Columns.Add(col_分值, "分值", 50, True)
        
        Set objCol = .Columns.Add(col_病人Id, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_反馈ID, "反馈ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_相关ID, "相关ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_对象ID, "对象ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_子文档ID, "子文档ID", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的反馈..."
            .ShadeGroupHeadings = True
        End With
        .ShowGroupBox = True
        .ShowItemsInGroups = False '是否按排序自分理处分组
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(col_状态)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_状态)
        .SortOrder(0).SortAscending = True
                
        .SortOrder.Add .Columns(col_处理人) '处理人为空的，用于区分暂存和待处理
        .SortOrder(1).SortAscending = True
        
        .SortOrder.Add .Columns(col_病人Id)
        .SortOrder(2).SortAscending = True
        
        .SortOrder.Add .Columns(col_反馈时间)
        .SortOrder(3).SortAscending = False
        
    End With
End Sub

Private Sub picData_Resize()
    rptData.Left = 0: rptData.Top = 0
    rptData.Width = picData.ScaleWidth
    rptData.Height = picData.ScaleHeight
End Sub

Private Sub picPati_GotFocus()
    Call txtPati_GotFocus
End Sub

Private Sub rptData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHit As ReportHitTestInfo
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim objCol As ReportColumn, lngCount As Long
    
    If Button = 2 Then
        Set objHit = rptData.HitTest(X, Y)
        If objHit.ht = xtpHitTestHeader Then
            Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                lngCount = 1
                For Each objCol In rptData.Columns
                    If objCol.AllowRemove And objCol.Width > 0 Then
                        Set objControl = .Add(xtpControlButton, conMenu_Col + lngCount, objCol.Caption)
                        objControl.Parameter = objCol.ItemIndex
                        lngCount = lngCount + 1
                    End If
                Next
            End With
            objPopup.ShowPopup
        ElseIf objHit.ht = xtpHitTestReportArea Then
            If Not objHit.Row Is Nothing Then
                If objHit.Row.GroupRow Then
                    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
                    With objPopup.Controls
                        Set objControl = .Add(xtpControlButton, conMenu_AllCollapse, "折叠所有组")
                        Set objControl = .Add(xtpControlButton, conMenu_AllExpend, "展开所有组")
                        Set objControl = .Add(xtpControlButton, conMenu_CurCollapse, "折叠当前组")
                            objControl.BeginGroup = True
                        Set objControl = .Add(xtpControlButton, conMenu_CurExpend, "展开当前组")
                    End With
                    objPopup.ShowPopup
                End If
            End If
        End If
    End If
End Sub

Private Function RefreshData() As Boolean
'功能：根据当前设置的条件读取反馈数据
Dim rsTmp As ADODB.Recordset, strSQL As String, strReturn As String, rsEmr As New ADODB.Recordset, strSQLEmr As String
Dim strPatis As String, objListItem As ListItem, curDate As Date
Dim objRecord As ReportRecord, objItem As ReportRecordItem, objRow As ReportRow, i As Long
Dim lngPreID As Long, lngPreIdx As Long
    
    Screen.MousePointer = 11
    If lvwPati.Visible = True Then lvwPati.Visible = False
    lvwPati.ListItems.Clear
        
    On Error GoTo errH
    
    '反馈数据，未归档的全部，历史归档的以时间为准
    If mintDataType = 0 Then
        strSQL = " And 反馈对象 IN(1,2,5,6,7,8,9)" '医生涉及的对象
    ElseIf mintDataType = 1 Then
        strSQL = " And 反馈对象 IN(3,4)" '护士涉及的对象
    End If
    If mvarCond.未处理 Then
        strSQL = "Select ID, 相关id, 病人id, 主页id, 记录性质, 记录状态, 反馈对象, 文件id, 反馈意见, 反馈人, 反馈时间, 处理期限, 处理说明, 处理人, 处理时间, 分制, 分值,补充说明, 子文档id From 病案反馈记录 Where 记录状态=1 And Instr([3],记录性质)>0" & strSQL
    ElseIf mvarCond.已处理 Then
        strSQL = _
            " Select ID, 相关id, 病人id, 主页id, 记录性质, 记录状态, 反馈对象, 文件id, 反馈意见, 反馈人, 反馈时间, 处理期限, 处理说明, 处理人, 处理时间, 分制, 分值,补充说明, 子文档id From 病案反馈记录 Where 记录状态 In(2,3) And Instr([3],记录性质)>0 And 处理时间 Between [4] And [5]" & strSQL & _
            " Union ALL" & _
            " Select ID, 相关id, 病人id, 主页id, 记录性质, 记录状态, 反馈对象, 文件id, 反馈意见, 反馈人, 反馈时间, 处理期限, 处理说明, 处理人, 处理时间, 分制, 分值,补充说明, 子文档id From 病案反馈历史 Where 记录状态 In(2,3) And Instr([3],记录性质)>0 And 处理时间 Between [4] And [5]" & strSQL
    End If
    
    'SQL中不排序提高效率,ReportControl有排序处理
    strSQL = _
        " Select NVL(B.姓名,C.姓名) 姓名,NVL(B.性别,C.性别) 性别,NVL(B.年龄,C.年龄) 年龄,B.住院号,B.出院病床 as 床号,D.名称 as 部门," & _
        " A.ID as 反馈ID,A.相关ID,A.病人ID,A.主页ID,A.记录性质,A.记录状态," & _
        " A.反馈对象,A.文件ID,A.子文档id,E.病历名称 as 住院病历,F.名称 as 护理记录,A.反馈意见,A.补充说明," & _
        " A.反馈人,A.反馈时间,A.处理期限,A.处理说明,A.处理人,A.处理时间,Decode(A.分制,1,'不合格',A.分值) as 分值" & _
        " From 病历文件列表 F,电子病历记录 E,部门表 D,病人信息 C,病案主页 B,(" & strSQL & ") A" & _
        " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And B.病人ID=C.病人ID And decode(length(a.文件id),32,0,a.文件id)=E.ID(+) And E.文件ID=F.ID(+)" & _
        IIf(mintDeptType = 0, " And B.出院科室ID=D.ID", " And B.当前病区ID=D.ID") & _
        IIf(mintDeptType = 0, " And B.出院科室ID=[1]", " And B.当前病区ID=[1]") & _
        IIf(mintDataType = 0, IIf(InStr(mstrPrivs, "本科病人") > 0 Or InStr(mstrPrivs, "全院病人") > 0, "", " And B.住院医师=[2]"), "") & _
        IIf(mintDataType = 0, IIf(mblnICU And InStr(mstrPrivs, "全院病人") = 0, " And B.住院医师=[2]", ""), "") & _
        " Order by 病人ID,反馈时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, UserInfo.姓名, _
        IIf(mvarCond.随机抽查, "1", "") & IIf(mvarCond.提交审查, "2", ""), mvarCond.开始时间, mvarCond.结束时间)
        
    '记录现在选中的反馈
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow And rptData.SelectedRows(0).Childs.Count = 0 Then
            lngPreIdx = rptData.SelectedRows(0).Index '用于快速重新定位
            lngPreID = rptData.SelectedRows(0).Record(col_反馈ID).Value
        End If
    End If
    
    txtNote.MaxLength = rsTmp.Fields("处理说明").DefinedSize
    curDate = zlDatabase.Currentdate
    rptData.Records.DeleteAll
    Do While Not rsTmp.EOF
        
        If InStr(strPatis & ",", "," & rsTmp!病人ID & ",") = 0 Then
            If strPatis = "" Then
                Set objListItem = lvwPati.ListItems.Add(, "_0_0", "全部")
                objListItem.SubItems(pcol_拼音简码) = "QB"
                txtPati.Text = "全部"
                txtPati.Tag = txtPati.Text
            End If
            strPatis = strPatis & "," & rsTmp!病人ID
            
            Set objListItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, rsTmp!姓名)
            If lvwPati.Tag = "_" & rsTmp!病人ID & "_" & rsTmp!主页ID Then
                objListItem.Selected = True
            End If
            
            objListItem.SubItems(pcol_住院号) = Nvl(rsTmp!住院号)
            objListItem.SubItems(pcol_床号) = Nvl(rsTmp!床号)
            objListItem.SubItems(pcol_性别) = Nvl(rsTmp!性别)
            objListItem.SubItems(pcol_年龄) = Nvl(rsTmp!年龄)
            objListItem.SubItems(pcol_部门) = Nvl(rsTmp!部门)
            objListItem.SubItems(pcol_拼音简码) = ZLCommFun.SpellCode(rsTmp!姓名 & "※0")
        End If
        
        Set objRecord = Me.rptData.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!记录状态))
        objItem.Caption = Decode(rsTmp!记录状态, 1, IIf(IsNull(rsTmp!处理人), "等待处理", "处理暂存"), 2, "等待复查", 3, "结束")
        objItem.Value = Val(rsTmp!记录状态)
        objItem.Icon = img16.ListImages("处理状态_" & objItem.Caption).Index - 1

        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!姓名)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!住院号, " "))) '加" "是为了排序合适
        
        Set objItem = objRecord.AddItem(zlStr.Lpad(Nvl(rsTmp!床号), 10)) 'Value用于排序
        objItem.Caption = Nvl(rsTmp!床号, " ") '为空时会被Value替代

        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!性别)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!年龄)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!部门)))
        Set objItem = objRecord.AddItem(Val(rsTmp!反馈对象))
        objItem.Caption = Decode(rsTmp!反馈对象, 1, "住院医嘱", 2, "住院病历", 3, "护理病历", 4, "护理记录", 5, "病案首页", 6, "医嘱报告", 7, "疾病证明", 8, "知情文件", 9, "临床路径")
        objItem.Icon = img16.ListImages("对象_" & Decode(rsTmp!反馈对象, 1, "医嘱", 2, "病历", 3, "病历", 4, "护理", 5, "首页", 6, "报告", 7, "文件", 8, "文件", 9, "路径")).Index - 1
        
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!反馈意见)))
        If rsTmp!反馈对象 = 1 Then
            objItem.Caption = "住院医嘱"
        ElseIf rsTmp!反馈对象 = 5 Then
            objItem.Caption = "病案首页"
        ElseIf rsTmp!反馈对象 = 4 Then
            objItem.Caption = Nvl(rsTmp!护理记录)
        ElseIf rsTmp!反馈对象 = 9 Then
            objItem.Caption = "临床路径"
        ElseIf Len(Nvl(rsTmp!文件ID)) < 32 Then
            objItem.Caption = Nvl(rsTmp!住院病历)
        ElseIf Len(Nvl(rsTmp!文件ID)) = 32 And Not gobjEmr Is Nothing Then '新版病历
            strSQLEmr = "Select Nvl(b.Subdoc_Title, a.Title) 病历名称" & vbNewLine & _
                    "From Bz_Doc_Log A, Bz_Doc_Tasks B" & vbNewLine & _
                    "Where a.Id = Hextoraw(:fileid) And a.Id = b.Real_Doc_Id" & IIf(Nvl(rsTmp!子文档ID) = "", "", " And b.Subdoc_Id = :subdocid")
            strReturn = gobjEmr.OpenSQLRecordset(strSQLEmr, rsTmp!文件ID & "^16^fileid" & IIf(Nvl(rsTmp!子文档ID) = "", "", "|" & Nvl(rsTmp!子文档ID) & "^16^subdocid"), rsEmr)
            If strReturn = "" Then
			If rsEmr.EOF THEN
				objItem.Caption = "【原始病历已不存在】"
			ELSE
                objItem.Caption = Nvl(rsEmr!病历名称)
			END If
            End If
        End If
        objItem.Caption = IIf(objItem.Caption <> "", objItem.Caption & ":", "") & Nvl(rsTmp!反馈意见)
        If rsTmp!记录状态 = 1 Then objItem.Bold = True
        
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!补充说明)))
            objItem.Caption = "" & Nvl(rsTmp!补充说明)
        Set objItem = objRecord.AddItem(CStr(Format(Nvl(rsTmp!处理期限), "yyyy-MM-dd HH:mm")))
        If Not IsNull(rsTmp!处理期限) And rsTmp!记录状态 = 1 Then
            If curDate > rsTmp!处理期限 Then objItem.ForeColor = vbRed
        End If
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!反馈人)))
        Set objItem = objRecord.AddItem(CStr(Format(Nvl(rsTmp!反馈时间), "yyyy-MM-dd HH:mm")))
        
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!处理说明)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!处理人)))
        Set objItem = objRecord.AddItem(CStr(Format(Nvl(rsTmp!处理时间), "yyyy-MM-dd HH:mm")))
        Set objItem = objRecord.AddItem("" & rsTmp!分值)
        
        Set objItem = objRecord.AddItem(Val(rsTmp!病人ID))
        Set objItem = objRecord.AddItem(Val(rsTmp!主页ID))
        Set objItem = objRecord.AddItem(Val(rsTmp!反馈ID))
        Set objItem = objRecord.AddItem(Val(Nvl(rsTmp!相关ID, 0)))
        Set objItem = objRecord.AddItem(Nvl(rsTmp!文件ID, "0"))
        Set objItem = objRecord.AddItem(Nvl(rsTmp!子文档ID, ""))
        
        rsTmp.MoveNext
    Loop
    rptData.Populate
    
    '定位到之前选择的病人上
    If Not (lvwPati.Tag = "" Or lvwPati.Tag = "_0_0") Then
        If Not lvwPati.SelectedItem Is Nothing Then
            Call lvwPati_KeyPress(vbKeyReturn)
        End If
    End If
        
    If rptData.Rows.Count = 0 Then
        txtNote.Locked = True '之前先Lock以方便判断
        txtNote.BackColor = txtResponse.BackColor
        txtResponse.Text = "": txtNote.Text = ""
        Me.stbThis.Panels(2).Text = ""
    Else
        If lngPreID <> 0 Then
            '先快速定位
            If lngPreIdx <= rptData.Rows.Count - 1 Then
                If Not rptData.Rows(lngPreIdx).GroupRow And rptData.Rows(lngPreIdx).Childs.Count = 0 Then
                    If rptData.Rows(lngPreIdx).Record(col_反馈ID).Value = lngPreID Then
                        Set objRow = rptData.Rows(lngPreIdx)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To rptData.Rows.Count - 1
                    If Not rptData.Rows(i).GroupRow And rptData.Rows(i).Childs.Count = 0 Then
                        If rptData.Rows(i).Record(col_反馈ID).Value = lngPreID Then
                            Set objRow = rptData.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        '取第一个非分组行
        If objRow Is Nothing Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow And rptData.Rows(i).Childs.Count = 0 Then Set objRow = rptData.Rows(i): Exit For
            Next
        End If
        Set rptData.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
        
        '选择某个病人时，其他记录是隐藏的
        If lvwPati.Tag = "" Or lvwPati.Tag = "_0_0" Then
            Me.stbThis.Panels(2).Text = "共有 " & rptData.Records.Count & " 条反馈记录"
        Else
            Me.stbThis.Panels(2).Text = ""
        End If
    End If
        
    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim lngPatID As Long, lngPageID As Long, intObject As Integer, strObjectID As String, strSubdocID As String
    If InStr(mstrPrivs, "审查反馈处理") = 0 Then Exit Sub
    
    If Not Row.GroupRow And Row.Childs.Count = 0 Then
        If Row.Record(col_状态).Value = 1 Or Row.Record(col_状态).Value = 2 Then
            lngPatID = CLng(Row.Record(col_病人Id).Value)
            lngPageID = CLng(Row.Record(col_主页ID).Value)
            intObject = CInt(Row.Record(col_反馈对象).Value)
            strObjectID = CStr(Row.Record(col_对象ID).Value)
            If Len(strObjectID) = 32 Then
                strSubdocID = CStr(Row.Record(col_子文档ID).Value)
                If strSubdocID <> "" Then strObjectID = strObjectID & "|" & strSubdocID
            End If
            RaiseEvent OpenObject(lngPatID, lngPageID, intObject, strObjectID)
        End If
    End If
End Sub

Private Sub rptData_SelectionChanged()
    Dim blnData As Boolean, blnModi As Boolean
    
    '显示详细数据
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow And rptData.SelectedRows(0).Childs.Count = 0 Then
            blnData = True
        End If
    End If
    
    '可否修改处理说明
    If blnData And InStr(mstrPrivs, "审查反馈处理") > 0 Then
        With rptData.SelectedRows(0)
            If .Record(col_状态).Value = 1 Or .Record(col_状态).Value = 2 And .Record(col_处理人).Value = UserInfo.姓名 Then
                blnModi = True
            End If
        End With
    End If
    txtNote.Locked = Not blnModi '之前先Lock以方便判断
    txtNote.BackColor = IIf(blnModi, vbWindowBackground, txtResponse.BackColor)
    
    If blnData Then
        txtResponse.Text = rptData.SelectedRows(0).Record(col_反馈意见).Value
        txtNote.Text = rptData.SelectedRows(0).Record(col_处理说明).Value
    Else
        txtResponse.Text = "": txtNote.Text = ""
    End If
End Sub

Private Sub txtNote_Change()
    If Visible And Not txtNote.Locked And rptData.SelectedRows.Count > 0 And Not mblnEditing Then
        If txtNote.Text <> rptData.SelectedRows(0).Record(col_处理说明).Value Then
            mblnEditing = True
        End If
    End If
End Sub

Private Sub txtNote_GotFocus()
    txtNote.SelStart = 0
    txtNote.SelLength = Len(txtNote.Text)
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Function SaveAllPaseData() As Boolean
'功能：完成所有暂存的反馈处理
    Dim colsql As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, strRows As String
        
    With rptData
        For i = 0 To .Records.Count - 1
            If .Records(i)(col_状态).Value = 1 And .Records(i)(col_处理人).Value = UserInfo.姓名 Then
                strSQL = "Zl_病案反馈记录_Process(" & .Records(i)(col_反馈ID).Value & ",2,'" & .Records(i)(col_处理说明).Value & "',1)"
                colsql.Add strSQL, "C" & colsql.Count + 1
            End If
        Next
        
        If colsql.Count = 0 Then
            MsgBox "你没有暂存的反馈处理记录，无需进行完成操作。", vbInformation, gstrSysName
            Exit Function
                
        ElseIf MsgBox("确实要将你暂存的所有反馈处理进行完成操作吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
        
        On Error GoTo errH
        If colsql.Count > 0 Then
            gcnOracle.BeginTrans: blnTrans = True
                For i = 1 To colsql.Count
                    Call zlDatabase.ExecuteProcedure(colsql("C" & i), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            
            '刷新界面
            Call RefreshData
        End If
    End With
        
    SaveAllPaseData = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData(ByVal blnPauseSave As Boolean) As Boolean
'参数：blnPauseSave-True=暂存，-False完成当前行
    Dim curDate As Date
    Dim strSQL As String, blnDel As Boolean
       
    If rptData.SelectedRows.Count = 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Or rptData.SelectedRows(0).Childs.Count > 0 Then Exit Function
        
    
    With rptData.SelectedRows(0)
        '非编辑状态下可“完成单条”（处理说明未变）
        If .Record(col_处理说明).Value <> txtNote.Text Or mvarCond.未处理 And Not blnPauseSave Then
            '输入检查
            If ZLCommFun.ActualLen(txtNote.Text) > txtNote.MaxLength Then
                MsgBox "处理说明内容太长，最多允许 " & txtNote.MaxLength \ 2 & " 个汉字或 " & txtNote.MaxLength & " 个字符。", vbInformation, gstrSysName
                txtNote.SetFocus: Exit Function
            End If
            
            '输入确认
            curDate = zlDatabase.Currentdate
            If txtNote.Text <> "" Then
                strSQL = "Zl_病案反馈记录_Process(" & .Record(col_反馈ID).Value & ",2,'" & Replace(txtNote.Text, "'", "''") & "'," & IIf(blnPauseSave, 0, 1) & ")"
                On Error GoTo errH
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                On Error GoTo 0
                                
                If Not blnPauseSave And mvarCond.未处理 Then
                    blnDel = True
                Else
                    .Record(col_处理说明).Value = txtNote.Text
                    .Record(col_处理时间).Value = CStr(Format(curDate, "yyyy-MM-dd HH:mm"))
                    .Record(col_处理人).Value = UserInfo.姓名
                    .Record(col_反馈意见).Bold = False
                    .Record(col_处理期限).ForeColor = Me.ForeColor
                    .Record(col_状态).Value = IIf(blnPauseSave, 1, 2)
                    .Record(col_状态).Caption = IIf(blnPauseSave, "处理暂存", "等待复查")
                    .Record(col_状态).Icon = img16.ListImages("处理状态_" & IIf(blnPauseSave, "处理暂存", "等待复查")).Index - 1
                End If
            Else
                strSQL = "Zl_病案反馈记录_Process(" & .Record(col_反馈ID).Value & ",1)"
                On Error GoTo errH
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                On Error GoTo 0
                
                If mvarCond.已处理 Then
                    blnDel = True
                Else
                    .Record(col_处理说明).Value = Empty
                    .Record(col_处理时间).Value = Empty
                    .Record(col_处理人).Value = Empty
                    .Record(col_反馈意见).Bold = True
                    If .Record(col_处理期限).Value <> "" Then
                        If curDate > CDate(.Record(col_处理期限).Value) Then
                            .Record(col_处理期限).ForeColor = vbRed
                        End If
                    End If
                    .Record(col_状态).Value = 1
                    .Record(col_状态).Caption = "等待处理" '暂存的清空处理意见后仍是等待处理
                    .Record(col_状态).Icon = img16.ListImages("处理状态_等待处理").Index - 1
                End If
            End If
            
            mblnOK = True
        End If
    End With
    
    If mblnOK Then
        If blnDel Then
            Call RefreshData
        Else
            rptData.Populate '分组可能变了，所以不能用redraw
        End If
        Me.stbThis.Panels(2).Text = "操作成功"
    End If
    
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdPati_Click()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    If cmdPati.Tag = "打开" Then
        cmdPati.Tag = ""
        lvwPati.Visible = False
        If txtPati.Enabled And txtPati.Visible Then txtPati.SetFocus
    Else
        cmdPati.Tag = "打开"
        If lvwPati.Tag <> "" And lvwPati.Tag <> "_0_0" Then
            lvwPati.ListItems(lvwPati.Tag).Selected = True
            lvwPati.SelectedItem.EnsureVisible
        End If
        lvwPati.Left = lngLeft + picPati.Left
        lvwPati.Top = lngTop + txtPati.Top
        lvwPati.ZOrder
        lvwPati.Visible = True
        lvwPati.SetFocus
    End If
End Sub



Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim lng病人ID As Long, lng主页ID As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            lng病人ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(0))
            lng主页ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1))
            
            Call ExecFilterPati(lng病人ID, lng主页ID)
            
            txtPati.Text = lvwPati.SelectedItem.Text
            txtPati.Tag = txtPati.Text
            
            lvwPati.Tag = lvwPati.SelectedItem.Key
            cmdPati.Tag = ""
            lvwPati.Visible = False
        End If
    End If
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
    cmdPati.Tag = ""
End Sub

Private Sub txtPati_GotFocus()
    txtPati.SelStart = 0
    txtPati.SelLength = Len(txtPati.Text)
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtPati_Validate(False)
    End If
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    Dim objItem As ListItem
    Dim strInput As String, blnABC As Boolean, blnFind As Boolean
    
    strInput = UCase(txtPati.Text)
    If strInput <> lblPati.Tag Then
        blnABC = ZLCommFun.IsCharAlpha(strInput)
        
        For Each objItem In lvwPati.ListItems
            If objItem.Text <> "全部" Then
                If blnABC Then
                    If objItem.SubItems(pcol_拼音简码) <> "" Then
                        If strInput Like objItem.SubItems(pcol_拼音简码) & "*" Then blnFind = True
                    End If
                Else
                    If strInput Like objItem.Text & "*" Then blnFind = True
                End If
            End If
            
            If blnFind Then
                objItem.Selected = True
                Call lvwPati_KeyPress(vbKeyReturn)
                Exit For
            End If
        Next
        If blnFind = False Then txtPati.Text = txtPati.Tag
    End If
End Sub

Private Sub ExecFilterPati(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
'功能：按病人过滤反馈记录
'参数：lng病人ID = 0 And lng主页ID = 0指全部显示
    
    Dim i As Long
    
    For i = 0 To rptData.Records.Count - 1
        If lng病人ID = 0 And lng主页ID = 0 Then
            If rptData.Records(i).Visible = False Then rptData.Records(i).Visible = True
        Else
            If rptData.Records(i).Item(col_病人Id).Value = lng病人ID And rptData.Records(i).Item(col_主页ID).Value = lng主页ID Then
                rptData.Records(i).Visible = True
            Else
                rptData.Records(i).Visible = False
            End If
        End If
    Next
    rptData.Populate
    If rptData.Rows.Count > 0 Then Set rptData.FocusedRow = rptData.Rows(1)
End Sub
