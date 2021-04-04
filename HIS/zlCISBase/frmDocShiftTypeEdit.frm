VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmDocShiftTypeEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医生交接班病人类型-新增"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   Icon            =   "frmDocShiftTypeEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8505
   StartUpPosition =   1  '所有者中心
   Begin XtremeSyntaxEdit.SyntaxEdit synSQL 
      Height          =   2895
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   7215
      _Version        =   983043
      _ExtentX        =   12726
      _ExtentY        =   5106
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   0   'False
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "验证(&C)"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Frame fraLine2 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   5160
      Width           =   9375
   End
   Begin VB.Frame fraLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   9375
   End
   Begin VB.TextBox txtBegin 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "可使用变量[时间格式]"
      Top             =   525
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtSName 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5760
      TabIndex        =   5
      Top             =   5400
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   7
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Label lblxing 
      AutoSize        =   -1  'True
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5640
      TabIndex        =   15
      Top             =   585
      Width           =   90
   End
   Begin VB.Label lblDescript 
      AutoSize        =   -1  'True
      Caption         =   "可使用[项目名称]变量"
      Height          =   180
      Left            =   5760
      TabIndex        =   14
      Top             =   585
      Width           =   1800
   End
   Begin VB.Label lblExplain 
      Caption         =   $"frmDocShiftTypeEdit.frx":5C02
      Height          =   615
      Left            =   960
      TabIndex        =   13
      Top             =   4320
      Width           =   7215
   End
   Begin VB.Label lblSQL 
      AutoSize        =   -1  'True
      Caption         =   "提取SQL"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblBegin 
      AutoSize        =   -1  'True
      Caption         =   "起始描述"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   585
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "名称"
      Height          =   180
      Left            =   2640
      TabIndex        =   8
      Top             =   165
      Width           =   360
   End
   Begin VB.Label lblSName 
      AutoSize        =   -1  'True
      Caption         =   "简称"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   165
      Width           =   360
   End
End
Attribute VB_Name = "frmDocShiftTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrSName As String
Private mblnOK As Boolean

Public Function ShowMe(ByVal bytType As Byte, ByRef strSName As String) As Boolean
'bytType:1-新增；2-修改
    
    mbytType = bytType
    mstrSName = strSName
    Me.Show 1
    If mblnOK Then strSName = mstrSName
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    If CheckSQL Then
        MsgBox "验证成功！", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    
    If CheckData = False Then Exit Sub
    
    strSql = SynSQL.Text
    strSql = Replace(strSql, "'", "''")
    On Error GoTo errH
    gstrSql = "Zl_医生交接班病人类型_Edit(" & IIf(mbytType = 1, 1, 2) & ",'" & txtSName.Text & "','" & _
        mstrSName & "','" & txtName.Text & "','" & txtBegin.Text & "','" & SynSQL.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    mblnOK = True
    mstrSName = txtSName.Text
    Unload Me
    Exit Sub
errH:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    
    With SynSQL
        '设置控件的显示颜色方案为：SQL
        .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
        .SyntaxScheme = GetSqlColor
    End With
    mblnOK = False
    Select Case mbytType
        Case 1
            Me.Caption = "医生交接班病人类型-新增"
        Case 2
            Me.Caption = "医生交接班病人类型-修改"
            txtSName.Text = mstrSName
            Set rsTemp = rsPatiType(mstrSName)
            If rsTemp.RecordCount = 1 Then
                txtName.Text = rsTemp!名称
                txtBegin.Text = rsTemp!起始描述 & ""
                SynSQL.Text = rsTemp!提取SQL & ""
            End If
    End Select
End Sub

Private Function CheckData() As Boolean
'保存前检查数据
    
    If txtSName.Text = "" Then
        MsgBox "简称不能为空，请检查！"
        Call zlcontrol.ControlSetFocus(txtSName)
        Exit Function
    ElseIf zlstr.ActualLen(txtSName.Text) > 10 Then
        MsgBox "简称不能超过5个汉字，请检查！"
        Call zlcontrol.ControlSetFocus(txtSName)
        Exit Function
    End If

    If txtName.Text = "" Then
        MsgBox "名称不能为空，请检查！"
        Call zlcontrol.ControlSetFocus(txtName)
        Exit Function
    ElseIf zlstr.ActualLen(txtName.Text) > 20 Then
        MsgBox "名称不能超过10个汉字，请检查！"
        Call zlcontrol.ControlSetFocus(txtName)
        Exit Function
    End If
    
    If zlstr.ActualLen(txtBegin.Text) > 50 Then
        MsgBox "起始描述不能超过25个汉字，请检查！"
        Call zlcontrol.ControlSetFocus(txtBegin)
        Exit Function
    End If
    
    If Trim(SynSQL.Text) <> "" Then
        If CheckSQL = False Then Exit Function
    End If
    CheckData = True
End Function

Private Function CheckSQL() As Boolean
'校验SQL的正确性
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
        
    strSql = Trim(UCase(SynSQL.Text))
    If Trim(SynSQL.Text) = "" Then
        MsgBox "提取SQL不能为空，请检查！", vbInformation, "验证SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    ElseIf zlstr.ActualLen(strSql) > 4000 Then
        MsgBox "提取SQL不能超过4000字符，请检查！", vbInformation, "验证SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    strSql = Replace(strSql, " ", "")
    If InStr(strSql, "A.病人ID<>-1ANDA.主页ID<>-1") = 0 And InStr(strSql, "A.主页ID<>-1ANDA.病人ID<>-1") = 0 Then
        MsgBox "提取SQL中必须包含[a.病人ID<>-1 And a.主页ID<>-1]条件"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    On Error GoTo errH
    gstrSql = "Select 提取sql From 医生交接班病人类型 Where 简称 <> [1] And 提取SQL is not null order by 顺序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "获取病人类型信息", mstrSName)
    If rsTemp.RecordCount > 0 Then
        strSql = rsTemp!提取SQL
        strSql = SynSQL.Text & vbNewLine & "Union All " & strSql
        strSql = UCase(strSql)
        strSql = Replace(strSql, "[开始时间]", zlstr.To_Date(Now))
        strSql = Replace(strSql, "[结束时间]", zlstr.To_Date(Now))
        strSql = Replace(strSql, "[科室ID]", "[1]")
        On Error Resume Next
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取病人类型信息", 0)
        If err.Number = 0 Then
            CheckSQL = True
            Exit Function
        Else
            MsgBox "提取SQL书写不正确，请检查！" & vbNewLine & err.Description, vbInformation, "验证SQL"
            Call zlcontrol.ControlSetFocus(SynSQL)
            Exit Function
        End If
    Else
        '如果数据库中没有一条数据，则进行字段的检查
        strSql = Trim(UCase(SynSQL.Text))
        If Not strSql Like "SELECT*病人ID,主页ID,姓名,性别,年龄,床号,标识号,入院时间,入院方式,出院科室ID*" Then
            MsgBox "提取SQL书写不正确，请检查！", vbInformation, "验证SQL"
            Call zlcontrol.ControlSetFocus(SynSQL)
            Exit Function
        End If
    End If
    CheckSQL = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Function

Private Sub synSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        SynSQL.Paste
    ElseIf KeyCode = vbKeyZ And Shift = 2 Then
        SynSQL.Undo
    ElseIf KeyCode = vbKeyY And Shift = 2 Then
        SynSQL.Redo
    ElseIf KeyCode = vbKeyC And Shift = 2 Then
        SynSQL.Copy
    ElseIf KeyCode = vbKeyA And Shift = 2 Then
        SynSQL.SelectAll
    End If
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub txtSName_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub
