VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmIdentify贵阳黑名单登记 
   Caption         =   "贵阳医保黑名单登记"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "frmIdentify贵阳黑名单登记.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   8385
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra药品明细查询 
      Caption         =   "选择保险用户"
      Height          =   2955
      Left            =   165
      TabIndex        =   2
      Top             =   105
      Width           =   8055
      Begin VB.TextBox txt医保号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   1965
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1380
         Width           =   1965
      End
      Begin VB.TextBox txt身份证号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1830
         Width           =   1965
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1830
         Width           =   1965
      End
      Begin VB.TextBox txt人员类别 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1395
         Width           =   1965
      End
      Begin VB.TextBox txt卡号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   885
         Width           =   1965
      End
      Begin VB.CommandButton cmd选择 
         Caption         =   "…"
         Height          =   300
         Left            =   7380
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   255
      End
      Begin VB.TextBox txt单位编码 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2340
         Width           =   1965
      End
      Begin VB.TextBox txt单位名称 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2340
         Width           =   1965
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2565
         TabIndex        =   12
         Top             =   270
         Width           =   5085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病人ID、姓名、住院号(&I)"
         Height          =   180
         Left            =   255
         TabIndex        =   21
         Top             =   360
         Width           =   2070
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   10385
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   10385
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label lab医保号 
         AutoSize        =   -1  'True
         Caption         =   "医保号"
         Height          =   180
         Left            =   420
         TabIndex        =   20
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lab卡号 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Left            =   5055
         TabIndex        =   19
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lab人员类别 
         AutoSize        =   -1  'True
         Caption         =   "人员类别"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lab性别 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   600
         TabIndex        =   17
         Top             =   1890
         Width           =   360
      End
      Begin VB.Label lab姓名 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   5055
         TabIndex        =   16
         Top             =   1455
         Width           =   360
      End
      Begin VB.Label lab身份证号 
         AutoSize        =   -1  'True
         Caption         =   "身份证号"
         Height          =   180
         Left            =   4695
         TabIndex        =   15
         Top             =   1890
         Width           =   720
      End
      Begin VB.Label lab单位名称 
         AutoSize        =   -1  'True
         Caption         =   "单位名称"
         Height          =   180
         Left            =   4695
         TabIndex        =   14
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label lab单位编码 
         AutoSize        =   -1  'True
         Caption         =   "单位编码"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   5535
      TabIndex        =   1
      Top             =   5655
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   6945
      TabIndex        =   0
      Top             =   5655
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfProject 
      Height          =   2355
      Left            =   165
      TabIndex        =   22
      ToolTipText     =   "Shift+Delete删除当前行"
      Top             =   3120
      Width           =   8055
      _cx             =   14208
      _cy             =   4154
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIdentify贵阳黑名单登记.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbrDelete 
      Left            =   3375
      Top             =   5685
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "提示：请点击收费细目编码选择"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   5737
      Width           =   2520
   End
End
Attribute VB_Name = "frmIdentify贵阳黑名单登记"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytEditMode            As Byte             '判断是登记还是修改登记
Private mstr医保号              As String
Private mrsD                    As ADODB.Recordset
Private mintInsure              As Integer
Private mstrSortID              As String
Private mblnCancel              As Boolean

Dim rsTmp                       As ADODB.Recordset
Dim strSQL                      As String
Dim sngX                        As Single
Dim sngY                        As Single
Dim sngH                        As Single

Const strSickFields = "select a.医保号 as ID,a.医保号, a.卡号,a.人员身份 as 人员类别,b.姓名,b.性别,b.身份证号,b.合同单位id as 单位编码,b.工作单位 as 单位名称 from 保险帐户 a , 病人信息 b where a.病人ID = b.病人id And a.险类 =" & TYPE_贵阳市
Const strProjectFields = "select id,decode(类别,'5','西成药','6','中成药','7','中草药','其它类别') as 类别,id as 收费细目ID,编码 as 收费细目编码,名称 as 收费细目名称,规格 as 收费细目规格 from 收费细目  "

Public Property Get blnCancel()
    blnCancel = mblnCancel
End Property

Public Property Let bytEditMode(ByVal vEditMode As Byte)
    mbytEditMode = vEditMode
End Property

Public Property Let intinsure(ByVal vintInsure As Integer)
    mintInsure = vintInsure
End Property
 
Public Property Let str医保号(ByVal vstr医保号 As String)
    mstr医保号 = vstr医保号
End Property

Private Sub cbrDelete_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 0
            vsfProject_Delete
    End Select
End Sub

Private Sub cmdCancle_Click()
On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim strTableD       As String
    Dim strWhereD       As String
    Dim i               As Integer
    Dim sFileName       As String
    Dim blnTran         As Boolean
On Error GoTo ErrH
    blnTran = False
    If txt医保号.Text = "" Then
        MsgBox "未选择医保人员！", vbInformation, gstrSysName
        Exit Sub
    ElseIf vsfProject.Tag <> "TRUE" Then
        MsgBox "黑名单项目未更改！" & vbCrLf & "请点击取消", vbInformation, gstrSysName
        Exit Sub
    End If
    mstr医保号 = txt医保号.Text
    If mbytEditMode = 2 Then
        '记录修改前日志
        ' 表表名(用分号";"隔开)
        strTableD = "医保黑名单_贵阳;医保黑名单项目_贵阳"
        ' 表的条件(用分号";"隔开)
        strWhereD = "医保号='" & txt医保号.Text & "';医保号='" & txt医保号.Text & "'"
        ' 记录修改前的数据
        sFileName = EditFormerWriteFileA(strTableD, strWhereD)
    End If
    With gcnGYYB
        .BeginTrans
        blnTran = True
        '保存主表数据
        strSQL = "Zl_医保黑名单_贵阳_Update ('" & txt医保号.Text & "','" & txt卡号.Text & "','" & txt人员类别.Text & "','" & txt姓名.Text & "','" & txt性别.Text & "','" & txt身份证号.Text & "','" & txt单位编码.Text & "','" & txt单位名称.Text & "',1)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        strSQL = "Zl_医保黑名单项目_贵阳_Delete ('" & txt医保号.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        For i = 1 To vsfProject.Rows - 1
            If vsfProject.TextMatrix(i, vsfProject.ColIndex("收费细目ID")) <> "" Then
                strSQL = "Zl_医保黑名单项目_贵阳_Update ('" & txt医保号.Text & "','" & vsfProject.TextMatrix(i, vsfProject.ColIndex("收费细目ID")) & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
        blnTran = False
        .CommitTrans
    End With
    
    If mbytEditMode = 2 Then
        '记录修改后日志
        Call EditFormerWriteFileA(strTableD, strWhereD, sFileName)
        '保存修改日志
        AddLog "医保工具", "医保黑名单_贵阳", DBConnLTEdit, , sFileName, mstr医保号, , , "医保黑名单_贵阳", , True
    End If
    Unload Me
    Exit Sub
ErrH:
    If blnTran Then
        gcnGYYB.RollbackTrans
        gcnGYYB.Errors.Clear
    End If
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
    If mbytEditMode = 2 Then
        strSQL = strSickFields
        strSQL = strSQL & " And a.医保号 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr医保号)
        If Not ChkRsState(rsTmp) Then
            txt医保号.Text = Nvl(rsTmp!医保号)
            txt卡号.Text = Nvl(rsTmp!卡号)
            txt人员类别.Text = Nvl(rsTmp!人员类别)
            txt姓名.Text = Nvl(rsTmp!姓名)
            txt性别.Text = Nvl(rsTmp!性别)
            txt身份证号.Text = Nvl(rsTmp!身份证号)
            txt单位编码.Text = Nvl(rsTmp!单位编码)
            txt单位名称.Text = Nvl(rsTmp!单位名称)
        End If
        txtFind.Locked = True
        txtFind.BackColor = &H80000000
        cmd选择.Enabled = False
    End If
    Call dDataload
    '
    With cbrDelete.KeyBindings
        .Add 4, vbKeyDelete, 0                 'Shift +Delete
        .Add 0, vbKeyDelete, 0                 'Shift +Delete
    End With
    
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub dDataload()
    On Error GoTo ErrH
    strSQL = "SELECT A.收费细目ID,类别,收费细目编码,收费细目名称,收费细目规格 " & vbCrLf & _
            "FROM 医保黑名单项目_贵阳 A ,(SELECT DECODE(类别,'5','西成药','6','中成药','其他类别') AS 类别, ID AS 收费细目ID, 编码 AS 收费细目编码,名称 AS 收费细目名称,规格 AS 收费细目规格 FROM 收费细目 ) B" & vbCrLf & _
            "WHERE A.收费细目ID = B.收费细目ID"
    strSQL = strSQL & vbCrLf & "And A.医保号 =[1]   "
    Set mrsD = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr医保号)
    Set vsfProject.DataSource = mrsD
    vsfProject.Rows = vsfProject.Rows + 1
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub cmd选择_Click()
    strSQL = strSickFields & " And A.医保号 not in (Select 医保号 From 医保黑名单项目_贵阳)"
    Call SickSelect(strSQL)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode <> 13 Then Exit Sub
    Dim strCode As String, strWhere As String
    '周玉强调整过提取方式
    strCode = txtFind.Text
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then
        '病人ID
        strWhere = " And A.病人ID=" & Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then
        '住院号
        strWhere = " And b.住院号='" & Mid(strCode, 2) & "'"
    Else
        '医保号
        strWhere = " And (b.姓名 Like '%" & strCode & "%' or A.医保号 like '%" & strCode & "%')"
    End If
    strSQL = strSickFields & " And A.医保号 not in (Select 医保号 From 医保黑名单项目_贵阳) " & strWhere
    Call SickSelect(strSQL)
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub SickSelect(sSql As String)
    Dim vRect       As RECT
    On Error GoTo ErrH
    vRect = GetControlRect(txtFind.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txtFind.Height
    strSQL = sSql
    Set rsTmp = zlDatabase.ShowSQLSelect( _
            Nothing, strSQL, 0, "医保病种选择", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, mintInsure, txtFind.Text _
            )
    If Not ChkRsState(rsTmp) Then
        txt医保号.Text = Nvl(rsTmp!医保号)
        txt卡号.Text = Nvl(rsTmp!卡号)
        txt人员类别.Text = Nvl(rsTmp!人员类别)
        txt姓名.Text = Nvl(rsTmp!姓名)
        txt性别.Text = Nvl(rsTmp!性别)
        txt身份证号.Text = Nvl(rsTmp!身份证号)
        txt单位编码.Text = Nvl(rsTmp!单位编码)
        txt单位名称.Text = Nvl(rsTmp!单位名称)
    Else
        MsgBox "没有找到病人信息!", vbInformation, gstrSysName
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
 
Private Sub vsfProject_AfterEdit(ByVal Row As Long, ByVal COL As Long)
    On Error GoTo ErrH
    vsfProject.Tag = "TRUE"
    Call vsfProject_KeyPressEdit(Row, COL, 13)
    If COL = vsfProject.ColIndex("收费细目编码") Then

        vsfProject.EditText = UCase(vsfProject.EditText)
        strSQL = strProjectFields & "  where (编码 like '%' || [1] || '%' or 名称  like '%' || [1] || '%' or zlSpellCode(名称)  like '%' || [1] || '%')"
        Call CalcPosition(sngX, sngY, vsfProject)
        sngY = sngY - vsfProject.CellHeight
        sngH = vsfProject.CellHeight
        DoEvents
        If Trim(vsfProject.EditText) = "" Then
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目ID")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目编码")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目名称")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目规格")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("类别")) = ""
            Exit Sub
        End If
        Set rsTmp = zlDatabase.ShowSQLSelect( _
                Nothing, strSQL, 0, "收费细目病种选择", False, _
                "", "", False, False, True, _
                sngX, sngY, sngH, False, False, _
                False, vsfProject.EditText _
                )
        If ChkRsState(rsTmp) Then
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目ID")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目编码")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目名称")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目规格")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("类别")) = ""
        Else
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目ID")) = rsTmp!ID
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目编码")) = rsTmp!收费细目编码
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目名称")) = rsTmp!收费细目名称
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("类别")) = rsTmp!类别
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目规格")) = rsTmp!收费细目规格
            If vsfProject.Rows - 1 = vsfProject.Row Then vsfProject.Rows = vsfProject.Rows + 1
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    On Error GoTo ErrH
    With vsfProject
        Select Case COL
            Case .ColIndex("收费细目编码")
                vsfProject.ComboList = "|..."
            Case Else
                .ComboList = ""
                Cancel = True
        End Select
        
    End With
    Exit Sub
ErrH:
    Err.Clear
    
End Sub

'==============================================================================
'=功能： 排序后定位记录 vsfProject
'==============================================================================
Private Sub vsfProject_AfterSort(ByVal COL As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
    lngRow = vsfProject.FindRow(mstrSortID, -1, vsfProject.ColIndex("疾病ID"), False, True)
    If lngRow > 0 Then vsfProject.Row = lngRow
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_BeforeMoveColumn(ByVal COL As Long, Position As Long)
    If COL = vsfProject.ColIndex("疾病ID") Then
        Position = -1
    Else
        If Position <= vsfProject.ColIndex("疾病ID") Then Position = COL
    End If
End Sub

'==============================================================================
'=功能： 某列不能拖动大小 vsfProject[图标]
'==============================================================================
Private Sub vsfProject_BeforeUserResize(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If COL = vsfProject.ColIndex("疾病ID") Then Cancel = True
End Sub

'==============================================================================
'=功能： 排序前记录ID vsfProject
'==============================================================================
Private Sub vsfProject_BeforeSort(ByVal COL As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("疾病ID"))
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_CellButtonClick(ByVal Row As Long, ByVal COL As Long)
    On Error GoTo ErrH
    vsfProject.Tag = "TRUE"
    If vsfProject.ColIndex("收费细目编码") = COL Then
        strSQL = strProjectFields
             
        Call CalcPosition(sngX, sngY, vsfProject)
        sngY = sngY - vsfProject.CellHeight
        sngH = vsfProject.CellHeight
        
        
        Set rsTmp = zlDatabase.ShowSQLSelect( _
                Nothing, strSQL, 0, "收费细目病种选择", False, _
                "", "", False, False, True, _
                sngX, sngY, sngH, False, False, _
                False, "" _
                )
        If ChkRsState(rsTmp) Then
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目ID")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目编码")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目名称")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目规格")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("类别")) = ""
        Else
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目ID")) = rsTmp!ID
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目编码")) = rsTmp!收费细目编码
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目名称")) = rsTmp!收费细目名称
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("类别")) = rsTmp!类别
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目规格")) = rsTmp!收费细目规格
            If vsfProject.Rows - 1 = vsfProject.Row Then vsfProject.Rows = vsfProject.Rows + 1
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_DblClick()
On Error GoTo ErrH
    If vsfProject.MouseRow = 0 Or vsfProject.MouseCol = 0 Then Exit Sub
    If vsfProject.TextMatrix(vsfProject.MouseRow, vsfProject.MouseCol) = "" Then Exit Sub
    Clipboard.SetText (vsfProject.TextMatrix(vsfProject.MouseRow, vsfProject.MouseCol))
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If vsfProject.ColIndex("收费细目编码") = vsfProject.COL Then
        '空格编辑
        If KeyAscii = vbKeySpace Then
            'KeyAscii = 39
            KeyAscii = 0
            SendKeys "{f2}"
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_KeyPressEdit(ByVal Row As Long, ByVal COL As Long, KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = asc("'") Then
       KeyAscii = 0
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Function vsfProject_Delete() As Long
    
    If vsfProject.Row = 1 And vsfProject.Rows = 2 Then
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目ID")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目编码")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目名称")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("收费细目规格")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("类别")) = ""
        Exit Function
    End If
    vsfProject.RemoveItem vsfProject.Row

End Function


