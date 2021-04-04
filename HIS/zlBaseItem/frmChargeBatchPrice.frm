VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeBatchPrice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "收费项目批量调价"
   ClientHeight    =   3525
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5685
   ClipControls    =   0   'False
   Icon            =   "frmChargeBatchPrice.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cbo价格等级 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2535
   End
   Begin VB.TextBox txtChargeType 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   885
      Width           =   2535
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   260
      Left            =   3840
      TabIndex        =   2
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox txtType 
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   525
      Width           =   2535
   End
   Begin VB.Frame fra调价方式 
      Caption         =   "调价"
      Height          =   2250
      Left            =   330
      TabIndex        =   15
      Top             =   1230
      Width           =   3795
      Begin VB.CheckBox chkByBase 
         Caption         =   "价格等级未设置价格时，在缺省价格的原价基础上调整(&P)"
         Height          =   375
         Left            =   270
         TabIndex        =   20
         Top             =   1830
         Width           =   3360
      End
      Begin VB.CheckBox chk子级 
         Caption         =   "包括该分类下所有子分类的项目(&G)"
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   1545
         Width           =   3195
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2310
         TabIndex        =   7
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   2310
         TabIndex        =   5
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optAdjust 
         Caption         =   "在原价基础上调整(&B)"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   750
         Width           =   2025
      End
      Begin VB.OptionButton optAdjust 
         Caption         =   "在原价基础上调整(&P)"
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Top             =   1140
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   172163075
         CurrentDate     =   36444
         MaxDate         =   401768
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "执行日期(&E)"
         Height          =   180
         Index           =   15
         Left            =   300
         TabIndex        =   18
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label5 
         Caption         =   "元"
         Height          =   180
         Left            =   3240
         TabIndex        =   17
         Top             =   810
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4350
      TabIndex        =   11
      Tag             =   "分类"
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4350
      TabIndex        =   10
      Tag             =   "分类"
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4350
      TabIndex        =   12
      Tag             =   "分类"
      Top             =   2820
      Width           =   1100
   End
   Begin VB.Label lbl价格等级 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费价格等级："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   19
      Top             =   173
      Width           =   1320
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费项目分类："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   14
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label lbl类别 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费项目类别："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   330
      TabIndex        =   13
      Top             =   555
      Width           =   1320
   End
End
Attribute VB_Name = "frmChargeBatchPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public datSingle As Date '单个分类下的最大日期
Public datAll As Date    '分类下所有项目的最大日期
Public dblSingle As Double   '单个分类下的最小金额
Public dblAll As Double      '分类下所有项目的最小金额
Public mblnCanUpdateAll As Boolean '是否允许操作所有项目：未启用价格等级或启用了价格等级有“所有院区”权限

Private Sub cbo价格等级_Click()
    chkByBase.Enabled = Not (cbo价格等级.Text = "缺省")
    chkByBase.value = vbUnchecked
End Sub

Private Sub chk子级_Click()
    If chk子级.value = 1 Then
        dtpBegin.MinDate = datAll
    Else
        dtpBegin.MinDate = datSingle
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsValid() As Boolean
'判断合法值
    If cbo价格等级.ListIndex = -1 Then
        MsgBox "请选择价格等级！", vbExclamation, gstrSysName
        If cbo价格等级.Visible And cbo价格等级.Enabled Then cbo价格等级.SetFocus
        Exit Function
    End If
    If optAdjust(0).value = True Then
        If IsNumeric(txtEdit(0).Text) = False Then
            MsgBox "请输入一个数值。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Val(txtEdit(0).Text) = 0 Then
            MsgBox "比例值不能为零。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Val(txtEdit(0).Text) <= -100 Then
            MsgBox "比例值不能低于-100%。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Val(txtEdit(0).Text) > 9999 Then
            MsgBox "比例值太大了。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            txtEdit(0).SetFocus
            Exit Function
        End If
    Else
        If IsNumeric(txtEdit(1).Text) = False Then
            MsgBox "请输入一个数值。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Val(txtEdit(1).Text) = 0 Then
            MsgBox "调整值不能为零。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Val(txtEdit(1).Text) + IIF(chk子级.value = 0, dblSingle, dblAll) <= 0 Then
            MsgBox "调整值至少要大于-" & IIF(chk子级.value = 0, dblSingle, dblAll) & "。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Val(txtEdit(1).Text) > 9999999 Then
            MsgBox "调整值太大了。", vbExclamation, gstrSysName
            zlControl.TxtSelAll txtEdit(1)
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    IsValid = True
End Function

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim int调整类型 As Integer '取值为1、按比例，本范围；2、按比例，全范围；3、按值，本范围；4、按值，全范围；
    Dim dbl调整值   As Double
    Dim str价格等级 As String
    
    If IsValid = False Then Exit Sub
    If MsgBox("批量调价会影响多个项目的价格，" & vbCrLf & "你确认已正确设置？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errMass
    
    If chk子级.value = 0 Then
        int调整类型 = IIF(optAdjust(0).value = True, 1, 3)
    Else
        int调整类型 = IIF(optAdjust(0).value = True, 2, 4)
    End If
    
    If cbo价格等级.Text = "所有价格等级" Then
        str价格等级 = "'所有'"
    ElseIf cbo价格等级.Text = "缺省" Then
        str价格等级 = "NULL"
    Else
        str价格等级 = "'" & cbo价格等级.Text & "'"
    End If
    
    'Zl_收费细目_Raisemass(
    gstrSQL = "zl_收费细目_RaiseMass("
    '  调整类型_In In Number,
    gstrSQL = gstrSQL & "" & int调整类型 & ","
    '  调整值_In   In Number,
    gstrSQL = gstrSQL & "" & IIF(optAdjust(0).value = True, Val(txtEdit(0).Text) / 100, Val(txtEdit(1).Text)) & ","
    '  执行日期_In In Date,
    gstrSQL = gstrSQL & "" & "to_date('" & Format(dtpBegin.value, "yyyy-MM-dd") & "','YYYY-MM-DD')" & ","
    '  终止日期_In In Date,
    gstrSQL = gstrSQL & "" & "to_date('" & Format(dtpBegin.value - 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '  调价人_In   In Varchar2,
    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
    '  分类id_In   In 收费项目目录.分类id%Type := Null,
    gstrSQL = gstrSQL & "" & IIF(lbl分类.Tag = "" Or lbl分类.Tag = "0", "null", lbl分类.Tag) & ","
    '  类别_In     In 收费项目目录.类别%Type := Null,
    gstrSQL = gstrSQL & "'" & lbl类别.Tag & "',"
    '  价格等级_In In 收费价目.价格等级%Type := '所有'
    gstrSQL = gstrSQL & "" & str价格等级 & ","
    '  站点_In     In 收费项目目录.站点%Type := Null
    gstrSQL = gstrSQL & "" & IIF(mblnCanUpdateAll, "NULL", "'" & gstrNodeNo & "'") & ","
    '  以缺省价格调整_In Number := 0
    gstrSQL = gstrSQL & "" & IIF(chkByBase.value = vbChecked, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If Not frmChargeManage.lvwMain_S.SelectedItem Is Nothing Then
        frmChargeManage.FillItem frmChargeManage.lvwMain_S.SelectedItem.Key
    End If
    Unload Me
    Exit Sub
errMass:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
On Error GoTo ErrHandle
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strReturn As String
    
    With frmSelCur
        strSQL = "Select Null,'所有类别' From Dual Union All Select 编码,名称 From 收费项目类别 where not 编码 in ('5','6','7') "
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            strReturn = .ShowCurrSel(Me, rsTmp, "编码,800,0,2;类别,1500,0,2", "类别选择器", False, Me.lbl类别.Tag, 0)
            If Trim(strReturn) <> "" Then
                txtType.Text = Split(strReturn, ",")(1)
                Me.lbl类别.Tag = Split(strReturn, ",")(0)
            End If
        Else
            MsgBox "无任何可用的类别，请与系统管理员联系！", vbExclamation, gstrSysName
            txtType.Text = "无"
            Me.lbl类别.Tag = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If txtEdit(0).Enabled = True Then txtEdit(0).SetFocus
End Sub

Private Sub Form_Load()
    Dim blnEnabled As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    cbo价格等级.Clear
    cbo价格等级.AddItem "所有价格等级"
    cbo价格等级.AddItem "缺省"
    If mblnCanUpdateAll Then
        strSQL = "Select Distinct a.名称 As 价格等级" & vbNewLine & _
                " From 收费价格等级 A" & vbNewLine & _
                " Where Nvl(a.是否适用普通项目, 0) = 1" & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Else
        strSQL = "Select Distinct a.名称 As 价格等级" & vbNewLine & _
                " From 收费价格等级 A, 收费价格等级应用 B" & vbNewLine & _
                " Where a.名称 = b.价格等级 And b.站点 = [1] And Nvl(a.是否适用普通项目, 0) = 1" & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))"

    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取价格等级", gstrNodeNo)
    Do While Not rsTemp.EOF
        cbo价格等级.AddItem Nvl(rsTemp!价格等级)
        rsTemp.MoveNext
    Loop
    If cbo价格等级.ListCount > 0 Then cbo价格等级.ListIndex = 0
        
    blnEnabled = IsPriceGradeEnabled()
    cbo价格等级.Enabled = blnEnabled
    chkByBase.Enabled = blnEnabled
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optAdjust_Click(Index As Integer)
    Dim lngOther As Long
    
    lngOther = IIF(Index = 0, 1, 0)
    txtEdit(Index).Enabled = True
    txtEdit(Index).BackColor = &H80000005
    txtEdit(Index).SetFocus
    txtEdit(lngOther).Enabled = False
    txtEdit(lngOther).BackColor = &H8000000F
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub
