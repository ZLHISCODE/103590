VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIdentify贵阳黑名单 
   Caption         =   "贵阳医保黑名单管理"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15870
   Icon            =   "frmIdentify贵阳黑名单.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   15870
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7545
      Left            =   3495
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7545
      ScaleWidth      =   45
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1755
      Width           =   45
   End
   Begin VB.PictureBox picBut 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   5070
      ScaleHeight     =   1755
      ScaleWidth      =   10620
      TabIndex        =   1
      Top             =   8325
      Width           =   10620
      Begin VB.CommandButton cmd修改登记 
         Caption         =   "修改登记(&E)"
         Height          =   345
         Left            =   6120
         TabIndex        =   11
         Top             =   750
         Width           =   1200
      End
      Begin VB.CommandButton cmd日志 
         Caption         =   "日志(&L)"
         Height          =   345
         Left            =   3075
         TabIndex        =   10
         Top             =   765
         Width           =   1200
      End
      Begin VB.CommandButton cmb登记 
         Caption         =   "登记(&R)"
         Height          =   345
         Left            =   4545
         TabIndex        =   9
         Top             =   750
         Width           =   1200
      End
      Begin VB.CommandButton cmd退出 
         Caption         =   "退出(&Q)"
         Height          =   345
         Left            =   9105
         TabIndex        =   8
         Top             =   750
         Width           =   1200
      End
      Begin VB.CommandButton cmb取消 
         Caption         =   "禁用启用(&S)"
         Height          =   345
         Left            =   7635
         TabIndex        =   7
         Top             =   750
         Width           =   1200
      End
      Begin VB.CheckBox chk非黑名单 
         Caption         =   "非黑名单"
         Height          =   360
         Left            =   6060
         TabIndex        =   6
         Top             =   90
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.CheckBox chk黑名单 
         Caption         =   "黑名单"
         Height          =   300
         Left            =   4770
         TabIndex        =   5
         Top             =   120
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2220
         TabIndex        =   3
         Top             =   105
         Width           =   2130
      End
      Begin VB.Label lab疾病 
         AutoSize        =   -1  'True
         Caption         =   "卡号、医保号、姓名(&F)"
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   180
         Width           =   1890
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   15000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   15000
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label pbrBar 
         BackColor       =   &H80000008&
         Height          =   240
         Left            =   1710
         TabIndex        =   2
         Top             =   2145
         Visible         =   0   'False
         Width           =   10380
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSick 
      Height          =   7545
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   7845
      _cx             =   13838
      _cy             =   13309
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIdentify贵阳黑名单.frx":000C
      ScrollTrack     =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsfProject 
      Height          =   7005
      Left            =   8160
      TabIndex        =   13
      Top             =   225
      Width           =   7515
      _cx             =   13256
      _cy             =   12356
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
      FormatString    =   $"frmIdentify贵阳黑名单.frx":0199
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
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   1275
      X2              =   13795
      Y1              =   8070
      Y2              =   8070
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   1290
      X2              =   13795
      Y1              =   8100
      Y2              =   8100
   End
End
Attribute VB_Name = "frmIdentify贵阳黑名单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure              As Integer
Private mrsM                    As ADODB.Recordset
Private mrsD                    As ADODB.Recordset
Dim strSQL                      As String

Const strSickFields = "Select Decode(状态, 0, '禁用', '1', '启用') As 状态, 次数, 医保号, 卡号, 人员类别, 姓名, 性别, 身份证号, 单位编码, 单位名称, 登记时间 From 医保黑名单_贵阳"

Public Property Let intinsure(ByVal vintInsure As Integer)
    mintInsure = vintInsure
End Property

Private Sub chk非黑名单_Click()
    On Error GoTo ErrH
    
    If chk黑名单.Value = chk非黑名单.Value Then
        mrsM.Filter = ""
    ElseIf chk黑名单.Value = 1 Then
        mrsM.Filter = "状态='启用'"
    ElseIf chk非黑名单.Value = 1 Then
        mrsM.Filter = "状态='禁用'"
    End If
    Set vsfSick.DataSource = mrsM
    Call vsfSick_RowColChange
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub chk黑名单_Click()
    Call chk非黑名单_Click
End Sub

Private Sub cmb登记_Click()
On Error GoTo ErrH
    With frmIdentify贵阳黑名单登记
        .bytEditMode = 1
        .intinsure = mintInsure
        .Show vbModal
    End With
    Set frmIdentify贵阳黑名单登记 = Nothing
    
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmb取消_Click()
    Dim intStop         As Integer
    
On Error GoTo ErrH
    If vsfSick.Row < 1 Or vsfSick.COL < 1 Then Exit Sub
    intStop = IIf(CStr(vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("状态"))) = "禁用", 1, 0)
    strSQL = "Zl_医保黑名单_贵阳_Update ('" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号")) & "','" & Trim(vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("卡号"))) & " ','" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("人员类别")) & "','" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("姓名")) & "','" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("性别")) & "','" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("身份证号")) & "','" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("单位编码")) & "','" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("单位名称")) & "'," & intStop & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mrsM.Requery
    vsfSick.Clear
    Set vsfSick.DataSource = mrsM
    vsfSetRow vsfSick, vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号")), "医保号"
    
    AddLog "医保工具", "医保黑名单_贵阳", DBConnLTSping, IIf(intStop = 1, "【启用】", "【禁用】"), , "" & vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号")) & "", , , "医保黑名单_贵阳"
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd日志_Click()
On Error GoTo ErrH
    If vsfSick.Row < 1 Or vsfSick.COL < 1 Then Exit Sub
    With frm医保操作日志
        .str模块 = "医保工具"
        .str功能 = "医保黑名单_贵阳"
        .str主键1 = CStr(vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号")))
        .Show vbModal
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd退出_Click()
    Unload Me
End Sub

Private Sub cmd修改登记_Click()
On Error GoTo ErrH
    If vsfSick.Row < 1 Or vsfSick.COL < 1 Then Exit Sub
    If CStr(vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("状态"))) = "禁用" Then
        MsgBox "该医保号【" & CStr(vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号"))) & "】黑名单项目已被禁用，不能修改！", vbCritical, gstrSysName
        Exit Sub
    End If
    With frmIdentify贵阳黑名单登记
        .bytEditMode = 2
        .intinsure = mintInsure
        .str医保号 = CStr(vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号")))
        .Show vbModal
        If .blnCancel Then
            Set frmIdentify贵阳黑名单登记 = Nothing
            Exit Sub
        End If
    End With
    Set frmIdentify贵阳黑名单登记 = Nothing
    mrsD.Requery
    Set vsfProject.DataSource = mrsD
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
    Call mDataload
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
On Error Resume Next
 
    If Me.WindowState = 1 Then Exit Sub
    
    sngTop = 0
    sngBottom = ScaleHeight - IIf(picBut.Visible, picBut.Height, 0)
    
    vsfSick.Top = sngTop
    vsfSick.Height = IIf(sngBottom - vsfSick.Top > 0, sngBottom - vsfSick.Top, 0)
    vsfSick.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = vsfSick.Left + vsfSick.Width
    
    vsfProject.Top = sngTop
    vsfProject.Left = picSplitV.Left + 35
    vsfProject.Width = ScaleWidth - vsfProject.Left
    vsfProject.Height = picSplitV.Height
    
    picBut.Move ScaleWidth - picBut.Width - 800, vsfSick.Height + 400
    
    With Line1
        .Y1 = picBut.Top - 120
        .Y2 = .Y1
        .X1 = 0
        .X2 = ScaleWidth
    End With
    With Line2
        .Y1 = Line1.Y1 + 30
        .Y2 = .Y1
        .X1 = 0
        .X2 = ScaleWidth
    End With
End Sub

Private Sub mDataload(Optional strKey As String = "")
On Error GoTo ErrH
    strSQL = strSickFields
    If strKey <> "" Then
        
    End If
    Set mrsM = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mintInsure)
    Set vsfSick.DataSource = mrsM
    If vsfSick.Rows > 1 Then vsfSick.Row = 1
    Call vsfSick_RowColChange
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub dDataload(str医保号 As String)
    On Error GoTo ErrH
    strSQL = "SELECT A.收费细目ID,类别,收费细目编码,收费细目名称,收费细目规格 " & vbCrLf & _
            "FROM 医保黑名单项目_贵阳 A ,(SELECT DECODE(类别,'5','西成药','6','中成药','7','中草药','其他类别') AS 类别, ID AS 收费细目ID, 编码 AS 收费细目编码,名称 AS 收费细目名称,规格 AS 收费细目规格 FROM 收费细目 ) B" & vbCrLf & _
            "WHERE A.收费细目ID = B.收费细目ID"
    strSQL = strSQL & vbCrLf & "And A.医保号 = [1]"
    Set mrsD = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str医保号)
    Set vsfProject.DataSource = mrsD
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button <> 1 Then Exit Sub
    
    With picSplitV
        .Move .Left + x
    End With
    Me.vsfSick.Width = picSplitV.Left
    Call Form_Resize
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    If KeyCode <> 13 Then Exit Sub
    vsfSetRow vsfSick, txtFind.Text, "卡号,医保号,姓名"
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfSick_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    On Error GoTo ErrH
    Cancel = True
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    On Error GoTo ErrH
    Cancel = True
    Exit Sub
ErrH:
    Err.Clear
End Sub
    
Private Sub vsfSick_CellChanged(ByVal Row As Long, ByVal COL As Long)
    Call vsfSick_RowColChange
End Sub

Private Sub vsfSick_Click()
    Call vsfSick_RowColChange
End Sub

Private Sub vsfSick_RowColChange()
On Error GoTo ErrH
    If vsfSick.Row < 1 Or vsfSick.COL < 1 Then
        dDataload ""
    Else
        dDataload vsfSick.TextMatrix(vsfSick.Row, vsfSick.ColIndex("医保号"))
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
