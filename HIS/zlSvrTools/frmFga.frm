VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFga 
   BackColor       =   &H00FFFFFF&
   Caption         =   "对象审计管理"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmFga.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   380
      Left            =   2640
      ScaleHeight     =   375
      ScaleWidth      =   3855
      TabIndex        =   23
      Top             =   580
      Width           =   3855
      Begin VB.TextBox txtFilter 
         ForeColor       =   &H00C0C0C0&
         Height          =   350
         Left            =   960
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "通过表名或规则名称定位规则"
         Top             =   10
         Width           =   2895
      End
      Begin VB.Label lblFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "规则定位"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   95
         Width           =   720
      End
   End
   Begin VB.PictureBox pctOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   6240
      ScaleHeight     =   3195
      ScaleWidth      =   4095
      TabIndex        =   20
      Top             =   1200
      Width           =   4125
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&Y)…"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "停用(&D)…"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "删除(&X)…"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加规则(&A)"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "关于审计的说明"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   3855
      End
   End
   Begin VB.PictureBox pctFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   10335
      TabIndex        =   17
      Top             =   5040
      Width           =   10335
      Begin VB.TextBox txtTerminal 
         Height          =   350
         Left            =   7200
         TabIndex        =   11
         Top             =   390
         Width           =   1935
      End
      Begin VB.TextBox txtGroup 
         Height          =   350
         Left            =   4320
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtUser 
         Height          =   350
         Left            =   840
         TabIndex        =   9
         Top             =   390
         Width           =   2535
      End
      Begin VB.ComboBox cboRule 
         Height          =   300
         ItemData        =   "frmFga.frx":803A
         Left            =   840
         List            =   "frmFga.frx":804A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   25
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&R)"
         Height          =   350
         Left            =   9240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   345
         Left            =   4320
         TabIndex        =   7
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         Format          =   56164355
         CurrentDate     =   43063.3914583333
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7200
         TabIndex        =   8
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         Format          =   119341059
         CurrentDate     =   43063.3913773148
      End
      Begin VB.Label lblTerminal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "客户端"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6540
         TabIndex        =   27
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "部门"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   26
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblUser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "用户名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblRuleList 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "规则名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         TabIndex        =   22
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblEnd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "结束时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6360
         TabIndex        =   19
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblStart 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "开始时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3480
         TabIndex        =   18
         Top             =   90
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRule 
      Height          =   1095
      Left            =   960
      TabIndex        =   14
      Top             =   960
      Width           =   4815
      _cx             =   8493
      _cy             =   1931
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsfLog 
      Height          =   1935
      Left            =   960
      TabIndex        =   15
      Top             =   3000
      Width           =   7335
      _cx             =   12938
      _cy             =   3413
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.Label lblLog 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "审计日志"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   16
      Top             =   2760
      Width           =   720
   End
   Begin VB.Image imgMain 
      Height          =   450
      Left            =   240
      Picture         =   "frmFga.frx":806E
      Top             =   600
      Width           =   465
   End
   Begin VB.Label lblRule 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "审计规则"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   13
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对象审计管理"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmFga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSchema As String    '策略
Private mstrObject As String
Private mstrPolicy As String
Private Enum Color
    tipColor = &H80000010
    txtColor = &H80000012
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
End Function

Private Sub cboRule_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{tab}"
    End If
End Sub

Private Sub cmdAdd_Click()
    frmFgaEdit.ShowMe
    Call GetPolicy
End Sub

Private Sub cmdEdit_Click()
    Dim strSql As String
    Dim strPolicy As String, strSchema As String, strObject As String
    
    On Error GoTo errh
    
    With vsfRule
        If .Row < 1 Then
            MsgBox "未选中审计规则,无法删除."
            Exit Sub
        End If
        strPolicy = .TextMatrix(.Row, .ColIndex("规则名"))
        strSchema = .TextMatrix(.Row, .ColIndex("所有者"))
        strObject = .TextMatrix(.Row, .ColIndex("表名"))
    End With
    
    If MsgBox("是否删除审计规则:" & strPolicy, vbQuestion + vbYesNo, "删除确认") = vbYes Then
    
        strSql = "Declare" & vbNewLine & _
                        "Begin" & vbNewLine & _
                        "  Dbms_Fga.Drop_Policy(Object_Schema => '" & strSchema & "', Object_Name => '" & strObject & "', Policy_Name => '" & strPolicy & "');" & vbNewLine & _
                        "End;"
        gcnOracle.Execute strSql
        
        Call GetPolicy
    End If
    
    Exit Sub
errh:
    MsgBox err.Description
End Sub

Private Sub cmdFind_Click()
    
    If vsfRule.Rows = 1 Or vsfRule.Row < 1 Then
        MsgBox "请先选中一个规则."
        Exit Sub
    End If
    
    Me.MousePointer = vbArrowHourglass
    frmMDIMain.stbThis.Panels(2).Text = "正在加载数据..."
    Call GetLog
    Me.MousePointer = vbDefault
    frmMDIMain.stbThis.Panels(2).Text = ""
End Sub

Private Sub cmdRefresh_Click()
    Call GetPolicy
End Sub

Private Sub cmdStop_Click()
    Dim strSql As String
    Dim strPolicy As String, strSchema As String, strObject As String
    
    On Error GoTo errh
    
    With vsfRule
        If .Row < 1 Then
            MsgBox "未选中审计规则,无法修改."
            Exit Sub
        End If
        strPolicy = .TextMatrix(.Row, .ColIndex("规则名"))
        strSchema = .TextMatrix(.Row, .ColIndex("所有者"))
        strObject = .TextMatrix(.Row, .ColIndex("表名"))
    End With
    
    Select Case cmdStop.Caption
    Case "停用(&D)…"
        If MsgBox("是否停用审计规则:" & strPolicy, vbQuestion + vbYesNo, "停用确认") = vbYes Then
        
            strSql = "Declare" & vbNewLine & _
                            "Begin" & vbNewLine & _
                            "  Dbms_Fga.Disable_Policy(Object_Schema => '" & strSchema & "', Object_Name => '" & strObject & "', Policy_Name => '" & strPolicy & "');" & vbNewLine & _
                            "End;"
            gcnOracle.Execute strSql
            
            Call GetPolicy
        End If
    Case "启用(&D)…"
        If MsgBox("是否启用审计规则:" & strPolicy, vbQuestion + vbYesNo, "停用确认") = vbYes Then
        
            strSql = "Declare" & vbNewLine & _
                            "Begin" & vbNewLine & _
                            "  Dbms_Fga.Enable_Policy(Object_Schema => '" & strSchema & "', Object_Name => '" & strObject & "', Policy_Name => '" & strPolicy & "');" & vbNewLine & _
                            "End;"
            gcnOracle.Execute strSql
            
            Call GetPolicy
        End If
    End Select
    
    Exit Sub
errh:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    Dim strPolicyCol As String, strLogCol As String
    lblHelp.Caption = "基于对象层面，可以对某张表、字段制定不同的审计规则（如Select、Update等）。数据库会记录下符合审计规则的SQL语句，便于快速回放事务场景。" & vbNewLine & vbNewLine & _
                                    "注：制定对象审计规则并生效后，对相关业务操作的执行性能会有一定影响，同时占用大量空间，所以对于非必要规则，请及时停用或删除。"
    
    strPolicyCol = "所有者,1800,1;表名,1380,1;规则名,2410,1;有效性,1995,1;操作类型,1710,1;列,2500,1"
    Call InitTable(vsfRule, strPolicyCol): vsfRule.Rows = 1
    
    strLogCol = "序号,600,1;规则名,1500,1;客户端,2410,1;用户名,1200,1;姓名,1500,1;部门,1500,1;SQL语句,3250,1;绑定变量,1560,1;时间,2200,1"
    Call InitTable(vsfLog, strLogCol)
     vsfLog.Rows = 1: vsfLog.FixedCols = 1
    
    dtpStart.value = date - 1
    dtpEnd.value = date + 1
    Call GetPolicy
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
        
    vsfRule.Height = 3050
    vsfRule.Width = Me.ScaleWidth - pctOpt.Width - vsfRule.Left - 180
    
    pctOpt.Top = vsfRule.Top
    pctOpt.Left = vsfRule.Left + vsfRule.Width + 60
    pctOpt.Height = vsfRule.Height

    pctFind.Top = pctOpt.Height + pctOpt.Top + 280
    pctFind.Left = Me.ScaleWidth - pctFind.Width - 60
        
    lblLog.Top = pctFind.Height + pctFind.Top - lblLog.Height - 60
    vsfLog.Top = pctFind.Top + pctFind.Height + 60
    vsfLog.Height = Me.ScaleHeight - vsfLog.Top - 350
    vsfLog.Width = pctOpt.Left + pctOpt.Width - vsfLog.Left
    
    pctFilter.Left = vsfRule.Left + vsfRule.Width - pctFilter.Width
End Sub

Private Sub GetPolicy()
    '功能:获取策略
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Integer
    
    strSql = "Select a.Object_Schema, a.Object_Name, a.Policy_Name, a.Enabled," & vbNewLine & _
                    "       Decode(a.Sel, 'YES', 'Select,', '') || Decode(a.Ins, 'YES', 'Insert,', '') ||" & vbNewLine & _
                    "        Decode(a.Upd, 'YES', 'Update,', '') || Decode(a.Del, 'YES', 'Delete,', '') Operators," & vbNewLine & _
                    "       Nvl(f_List2str(Cast(Collect(b.Policy_Column) As t_Strlist), ',', 1),'全部列') Columns" & vbNewLine & _
                    "From Dba_Audit_Policies A, Dba_Audit_Policy_Columns B" & vbNewLine & _
                    "Where a.Object_Schema = b.Object_Schema(+) And a.Object_Name = b.Object_Name(+) And a.Policy_Name = b.Policy_Name(+)" & vbNewLine & _
                    "Group By a.Object_Schema, a.Object_Name, a.Policy_Name, a.Enabled, a.Sel, a.Ins, a.Upd, a.Del"
                    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "GetdPolicy")
    
    With vsfRule
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rsTmp.RecordCount + .FixedRows
        i = .FixedRows
        cboRule.Clear
        Do While Not rsTmp.EOF
            .TextMatrix(i, .ColIndex("所有者")) = rsTmp!Object_Schema & ""
            .TextMatrix(i, .ColIndex("表名")) = rsTmp!Object_Name & ""
            .TextMatrix(i, .ColIndex("规则名")) = rsTmp!Policy_Name & ""
            cboRule.addItem rsTmp!Policy_Name & ""
            .TextMatrix(i, .ColIndex("有效性")) = rsTmp!Enabled & ""
            .TextMatrix(i, .ColIndex("列")) = rsTmp!Columns & ""
            .TextMatrix(i, .ColIndex("操作类型")) = IIf(Nvl(rsTmp!Operators) <> "", Left(Nvl(rsTmp!Operators), Abs(InStrRev(Nvl(rsTmp!Operators), ",") - 1)), "") 'SQL拼接字符串时多拼了一个逗号,这里去掉
            i = i + 1
            rsTmp.MoveNext
        Loop
        .Redraw = flexRDDirect
        If .Rows > 1 Then .Select 1, 0
    End With
End Sub

Private Sub GetLog()
    '功能:获取策略对应的日志
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Integer
    Dim strSchema As String, strName As String, strPolicy As String
    
    With vsfRule
        strPolicy = cboRule.Text
        
        For i = 1 To .Rows - 1
            If strPolicy = .TextMatrix(i, .ColIndex("规则名")) Then
                strSchema = .TextMatrix(i, .ColIndex("所有者"))
                strName = .TextMatrix(i, .ColIndex("表名"))
                Exit For
            End If
        Next
    End With
    
    strSql = "Select a.Policy_Name,a.Userhost, a.Db_User, a.Sql_Text, a.Sql_Bind, a.Statement_Type,A.Timestamp ,c.姓名, d.名称 部门" & vbNewLine & _
                    "From (Select a.Object_Schema, a.Object_Name, a.Policy_Name, a.Timestamp, a.Userhost, a.Db_User," & vbNewLine & _
                    "              a.Sql_Text, a.Sql_Bind, a.Statement_Type" & vbNewLine & _
                    "       From Dba_Fga_Audit_Trail A" & vbNewLine & _
                    "       Where a.Object_Schema = [1] And a.Object_Name = [2] And a.Policy_Name = [3] And" & vbNewLine & _
                    "             Timestamp Between [4] And [5]) A, 上机人员表 B, 人员表 C, 部门表 D, 部门人员 E" & vbNewLine & _
                    "Where a.Db_User = b.用户名(+) And b.人员id = c.Id(+) And c.Id = e.人员id(+) And e.部门id = d.Id(+) and e.缺省 = 1 " & vbNewLine & _
                    IIf(txtUser.Text <> "", "And b.用户名=[6]", "") & vbNewLine & _
                    IIf(txtGroup.Text <> "", "And d.名称 =[7]", "") & vbNewLine & _
                    IIf(txtTerminal.Text <> "", "And a.Userhost=[8]", "") & vbNewLine & _
                    "Order By A.timestamp Desc"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "GetdLog", strSchema, strName, strPolicy, _
                                                                        CDate(Format(dtpStart.value, "yyyy-MM-dd hh:mm:ss")), CDate(Format(dtpEnd.value, "yyyy-MM-dd hh:mm:ss")), _
                                                                        UCase(txtUser.Text), txtGroup.Text, txtTerminal.Text)
    With vsfLog
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rsTmp.RecordCount + .FixedRows
        i = .FixedRows
        .ColAlignment(0) = flexAlignCenterCenter
        Do While Not rsTmp.EOF
            .TextMatrix(i, .ColIndex("序号")) = i
            .TextMatrix(i, .ColIndex("规则名")) = rsTmp!Policy_Name & ""
            .TextMatrix(i, .ColIndex("客户端")) = rsTmp!Userhost & ""
            .TextMatrix(i, .ColIndex("用户名")) = rsTmp!Db_User & ""
            .TextMatrix(i, .ColIndex("姓名")) = rsTmp!姓名 & ""
            .TextMatrix(i, .ColIndex("部门")) = rsTmp!部门 & ""
            .TextMatrix(i, .ColIndex("SQL语句")) = Replace(Trim(rsTmp!Sql_Text & ""), Chr(10), "")
            .TextMatrix(i, .ColIndex("绑定变量")) = rsTmp!Sql_Bind & ""
            .Cell(flexcpData, i, .ColIndex("SQL语句")) = rsTmp!Sql_Text & ""
            .Cell(flexcpData, i, .ColIndex("绑定变量")) = rsTmp!Sql_Bind & ""
            .TextMatrix(i, .ColIndex("时间")) = rsTmp!TimeStamp & ""
            i = i + 1
            rsTmp.MoveNext
        Loop
        .Redraw = flexRDDirect
        If .Rows > 1 Then .Select 1, 0
    End With
End Sub

Private Sub txtFilter_GotFocus()
    If txtFilter.Text = "通过表名或规则名称定位规则" Then
        txtFilter.Text = ""
        txtFilter.ForeColor = txtColor
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call GetRowPos(vsfRule, txtFilter.Text, "表名,规则名")
    End If
End Sub

Private Sub txtFilter_LostFocus()
    If txtFilter.Text = "" Then
        txtFilter.Text = "通过表名或规则名称定位规则"
        txtFilter.ForeColor = tipColor
    End If
End Sub

Private Sub txtGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{tab}"
    End If
End Sub

Private Sub txtTerminal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{tab}"
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{tab}"
    End If
End Sub

Private Sub vsfLog_DblClick()
    With vsfLog
        If .Row < 1 Then Exit Sub
        frmFgaMore.ShowMe .Cell(flexcpData, .Row, .ColIndex("SQL语句")), .Cell(flexcpData, .Row, .ColIndex("绑定变量"))
    End With
End Sub

Private Sub vsfRule_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    With vsfRule
        If .Redraw = flexRDNone Or .Rows < 2 Then Exit Sub
        
        cboRule.Text = .TextMatrix(.Row, .ColIndex("规则名"))
        
        If .TextMatrix(.Row, .ColIndex("有效性")) = "YES" Then
            cmdStop.Caption = "停用(&D)…"
        Else
            cmdStop.Caption = "启用(&D)…"
        End If
    End With
End Sub
