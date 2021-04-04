VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeSortEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "费别设置"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmChargeSortEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra服务 
      Caption         =   "服务对象"
      Height          =   645
      Left            =   3360
      TabIndex        =   11
      Top             =   150
      Width           =   2595
      Begin VB.OptionButton opt服务 
         Caption         =   "所有"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opt服务 
         Caption         =   "住院"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton opt服务 
         Caption         =   "门诊"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fra科室 
      Caption         =   "适用科室"
      Height          =   3345
      Left            =   3360
      TabIndex        =   15
      Top             =   960
      Width           =   2595
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   350
         Left            =   1320
         TabIndex        =   19
         Top             =   2820
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   350
         Left            =   180
         TabIndex        =   18
         Top             =   2820
         Width           =   1100
      End
      Begin VB.ListBox lst科室 
         Enabled         =   0   'False
         Height          =   2220
         Left            =   180
         TabIndex        =   17
         Top             =   540
         Width           =   2235
      End
      Begin VB.CheckBox chk科室 
         Caption         =   "所有科室(&L)"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Value           =   1  'Checked
         Width           =   1305
      End
   End
   Begin VB.Frame frm基本 
      Caption         =   "基本情况"
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3150
      Begin VB.CheckBox chk缺省 
         Caption         =   "缺省(&E)"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   180
         TabIndex        =   30
         Top             =   3240
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1560
         TabIndex        =   28
         Top             =   1440
         Width           =   1450
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   100663299
         CurrentDate     =   40871
      End
      Begin VB.CheckBox chk初诊 
         Caption         =   "仅限初诊(&F)"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   2520
         Width           =   1485
      End
      Begin VB.OptionButton opt属性 
         Caption         =   "动态性项目(&Y)"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   2880
         Width           =   1485
      End
      Begin VB.OptionButton opt属性 
         Caption         =   "身份唯一项目(&I)"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   2220
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   900
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "简码"
         Top             =   1026
         Width           =   2055
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   900
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "名称"
         Top             =   648
         Width           =   2055
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   900
         MaxLength       =   2
         TabIndex        =   27
         Tag             =   "编码"
         Top             =   270
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1560
         TabIndex        =   29
         Top             =   1785
         Width           =   1450
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   100663299
         CurrentDate     =   40871
      End
      Begin VB.Label lbl缺省 
         Caption         =   "如果选中它，其它费别的本属性自动取消。"
         Height          =   420
         Left            =   180
         TabIndex        =   31
         Top             =   3600
         Width           =   2880
      End
      Begin VB.Label lbl属性 
         AutoSize        =   -1  'True
         Caption         =   "性质"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "有效结束日期(&P)"
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   1845
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "有效开始日期(&B)"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   1470
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   705
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.TextBox txtEdit 
      Height          =   750
      Index           =   3
      Left            =   1080
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   21
      Tag             =   "说明"
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -60
      TabIndex        =   22
      Top             =   5400
      Width           =   6270
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   210
      TabIndex        =   25
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3360
      TabIndex        =   23
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4590
      TabIndex        =   24
      Top             =   5640
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "说明(&X)"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   4560
      Width           =   630
   End
End
Attribute VB_Name = "frmChargeSortEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    text编码 = 0
    Text名称 = 1
    text简码 = 2
    Text说明 = 3
    Text开始 = 4
    Text结束 = 5
End Enum

Dim mstr名称 As String         '当前编辑费别的原名称
Dim mblnChange As Boolean      '是否改变了
Dim mBoundSelect As Integer    '当前服务范围

Private Sub cmdAdd_Click()
    Dim blnRe  As Boolean, lngIndex As Long
    Dim strID As String, str名称 As String, str编码 As String, str原有ID As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select distinct id,上级id,名称,编码 from 部门表 " & _
              "where 撤档时间=to_date('3000-01-01','YYYY-MM-DD') " & _
              "start with ID In " & _
              "(Select 部门ID From 部门性质说明 Where 服务对象 In (" & _
              Switch(opt服务(0).Value, "1,", opt服务(1).Value, "2,", opt服务(2).Value, "1,2,") & _
              "3)) connect by prior 上级id=ID Order By 编码"
    blnRe = frmTreeSel.ShowTree(gstrSQL, strID, str名称, str编码, str原有ID, "费别适用科室", "所有科室", False)
    
    If blnRe = True Then
        For lngIndex = 0 To lst科室.ListCount - 1
            If lst科室.ItemData(lngIndex) = Val(strID) Then
                '已经有该科室，不用再继续
                Exit Sub
            End If
        Next

        lst科室.AddItem "【" & str编码 & "】" & str名称
        lst科室.ItemData(lst科室.NewIndex) = strID
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    If lst科室.ListIndex < 0 Then Exit Sub
    lst科室.RemoveItem lst科室.ListIndex
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "会员等级设置"
        lbl属性.Caption = "会员等级属性"
        lblEdit(5).Caption = "等级说明(&X)"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save费别() = False Then Exit Sub
    
    Call frmChargeSortGrade.FillList
    If mstr名称 <> "" Then
        '修改成功
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    '连续新增
    For i = 1 To txtEdit.UBound
        txtEdit(i).Text = ""
    Next
    chk缺省.Value = 0
    txtEdit(text编码).Text = sys.MaxCode("费别", "编码", 2)
    txtEdit(text编码).SetFocus
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关费别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = txtEdit.LBound To txtEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(i)
            txtEdit(i).SetFocus
            Exit Function
        End If
    Next
    
    If Trim(txtEdit(text编码).Text) = "" Then
        MsgBox "编码不能为空。", vbInformation, gstrSysName
        txtEdit(text编码).Text = ""
        txtEdit(text编码).SetFocus
        Exit Function
    End If
    If Trim(txtEdit(Text名称).Text) = "" Then
        MsgBox "名称不能为空。", vbInformation, gstrSysName
        txtEdit(Text名称).Text = ""
        txtEdit(Text名称).SetFocus
        Exit Function
    End If
    
    If IsDate(dtpBegin.Value) And IsDate(dtpEnd.Value) Then
        If CDate(dtpBegin.Value) > CDate(dtpEnd.Value) Then
            MsgBox "有效期的开始日期不能大于结束日期。", vbInformation, gstrSysName
            dtpBegin.SetFocus
            Exit Function
        End If
    End If
    If chk科室.Value = 0 And lst科室.ListCount = 0 Then
        If MsgBox("本费别的适用科室不能所有科室，且又没选择指定科室。" & vbCrLf & "是否继续？", _
                    vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chk科室.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Save费别() As Boolean
'功能:保存编辑的内容到费别表中
'参数:
'返回值:成功返回True,否则为False
    On Error GoTo ErrHandle
    Dim lngCount As Long
    Dim str开始日期 As String, str结束日期 As String
    Dim str指定科室 As String
    
    str开始日期 = IIF(dtpBegin.Enabled = False Or IsDate(dtpBegin.Value) = False, _
                    "null", "to_date('" & dtpBegin.Value & "','YYYY-MM-dd')")
    str结束日期 = IIF(dtpEnd.Enabled = False Or IsDate(dtpEnd.Value) = False, _
                    "null", "to_date('" & dtpEnd.Value & "','YYYY-MM-dd')")
    If lst科室.Enabled = True Then
        For lngCount = 0 To lst科室.ListCount - 1
            str指定科室 = str指定科室 & lst科室.ItemData(lngCount) & ","
        Next
    End If
    
    gstrSQL = Trim(txtEdit(text编码).Text) & "','" & Trim(txtEdit(Text名称).Text) & "','" & _
            Trim(txtEdit(text简码).Text) & "','" & Trim(txtEdit(Text说明).Text) & "'," & _
            str开始日期 & "," & str结束日期 & "," & IIF(chk科室.Value = 1, 1, 2) & "," & _
            IIF(opt属性(0).Value = True, 1, 2) & "," & chk初诊.Value & "," & chk缺省.Value & ",'" & str指定科室 & "'," & _
            Switch(opt服务(0).Value, 1, opt服务(1).Value, 2, opt服务(2).Value, 3) & ")"
    If mstr名称 = "" Then       '新增一条记录
        gstrSQL = "zl_费别_Insert('" & gstrSQL
    Else    '修改
        gstrSQL = "zl_费别_update('" & mstr名称 & "','" & gstrSQL
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save费别 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 编辑费别(ByVal str名称 As String) As Boolean
'功能:用来与调用的费别管理窗口进行通讯的程序
'参数:str名称          当前编辑的费别名称
'返回值:编辑成功返回True,否则为False
    Dim rs费别 As New ADODB.Recordset
    Dim i As Integer
    
    mstr名称 = str名称
    
    On Error GoTo ErrHandle
    If str名称 <> "" Then
        rs费别.CursorLocation = adUseClient
        
        gstrSQL = "Select 编码, 名称, 简码, 有效开始, 有效结束, 适用科室, 属性, 仅限初诊, 缺省标志, 服务对象, 说明 From 费别 Where 名称 =[1] "
        Set rs费别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str名称)
                
        txtEdit(text编码).Text = rs费别("编码")
        txtEdit(Text名称).Text = rs费别("名称")
        txtEdit(text简码).Text = IIF(IsNull(rs费别("简码")), "", rs费别("简码"))
        If rs费别("有效开始") & "" <> "" Then
            dtpBegin.Value = CDate(Format(rs费别("有效开始"), "yyyy-MM-dd"))
        Else
            dtpBegin.Value = Null
        End If
        If rs费别("有效结束") & "" <> "" Then
            dtpEnd.Value = CDate(Format(rs费别("有效结束"), "yyyy-MM-dd"))
        Else
            dtpEnd.Value = Null
        End If
        txtEdit(Text说明).Text = IIF(IsNull(rs费别("说明")), "", rs费别("说明"))
        
        opt属性(IIF(rs费别("属性") = 2, 1, 0)).Value = True
        
        chk初诊.Value = IIF(rs费别("仅限初诊") = 1, 1, 0)
        opt服务(IIF(IsNull(rs费别("服务对象")), 2, Val(rs费别("服务对象")) - 1)).Value = True
        chk缺省.Value = IIF(rs费别("缺省标志") = 1, 1, 0)
        chk科室.Value = IIF(rs费别("适用科室") = 2, 0, 1) '2表示指定科室
        
        lst科室.Clear
        If chk科室.Value = 0 Then
            '读出指定科室
            gstrSQL = "select B.ID,B.编码,B.名称 from 费别适用科室 A,部门表 B " & _
                      " where A.费别=[1] and A.科室ID=B.ID order by B.编码"
            Set rs费别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr名称)
                        
            Do Until rs费别.EOF
                lst科室.AddItem "【" & rs费别("编码") & "】" & rs费别("名称")
                lst科室.ItemData(lst科室.NewIndex) = rs费别("ID")
                
                rs费别.MoveNext
            Loop
        End If
    Else
        txtEdit(text编码).Text = sys.MaxCode("费别", "编码", 2)
        dtpBegin.Value = Null
        dtpEnd.Value = Null
    End If
    Call SetEnable
    
    mblnChange = False
    
    For i = 0 To Me.opt服务.Count - 1
        If Me.opt服务(i).Value = True Then
            mBoundSelect = i
            Exit For
        End If
    Next
        
    
    frmChargeSortEdit.Show vbModal
    编辑费别 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub opt服务_Click(Index As Integer)
    '跟上次一样时退出
    If mBoundSelect = Index Then Exit Sub
    If Index = 2 Then
        mBoundSelect = 2
        Exit Sub
    End If
    If Me.lst科室.ListCount > 0 Then
        If MsgBox("您选择了另一个服务对象，现有选中的科室将被清除！是否继续？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Me.lst科室.Clear
            mBoundSelect = Index
        Else
            Me.opt服务(mBoundSelect).Value = True
            Me.opt服务(mBoundSelect).SetFocus
        End If
    End If
End Sub

Private Sub opt服务_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text名称 Then
        txtEdit(text简码).Text = zlStr.GetCodeByVB(txtEdit(Text名称).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = Text名称 Or Index = Text说明 Then
        OS.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Dim strDate As String
    
    If Index = Text开始 Or Index = Text结束 Then
        '处理日期
        strDate = zlCommFun.AddDate(txtEdit(Index).Text)
        If IsDate(strDate) Then
            txtEdit(Index).Text = Format(CDate(strDate), "yyyy-MM-dd")
        Else
            txtEdit(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          KeyAscii = 0
          SendKeys "{TAB}"
    ElseIf KeyAscii = Asc(":") Or KeyAscii = Asc(",") Then  '计算实收金额的函数Zl_Actualmoney返回的串用到了:号分隔符
        KeyAscii = 0
    End If
End Sub

Private Sub chk缺省_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chk缺省_Click()
    Call SetEnable
End Sub

Private Sub chk初诊_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chk初诊_Click()
    Call SetEnable
End Sub

Private Sub chk科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chk科室_Click()
    Call SetEnable
End Sub

Private Sub opt属性_Click(Index As Integer)
    Call SetEnable
End Sub

Private Sub opt属性_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lst科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub SetEnable()
'根据控件取值的不同，设置相关控件的可用性
    Dim bln普通 As Boolean
    
    mblnChange = True
    If chk缺省.Value = 1 Then
        bln普通 = False
        '只能使用特定值
        chk科室.Value = 1
        opt属性(0).Value = True
        chk初诊.Value = 0
    Else
        bln普通 = True
    End If
    lblEdit(Text开始).Enabled = bln普通
    lblEdit(Text结束).Enabled = bln普通
    dtpBegin.Enabled = bln普通
    dtpEnd.Enabled = bln普通
    opt属性(0).Enabled = bln普通
    opt属性(1).Enabled = bln普通
    chk科室.Enabled = bln普通
    
    If opt属性(1).Value = True Then
        chk初诊.Value = 0
    End If
    chk初诊.Enabled = bln普通 And opt属性(0).Value
    
    lst科室.Enabled = (chk科室.Value = 0)
    cmdAdd.Enabled = lst科室.Enabled
    cmdDelete.Enabled = lst科室.Enabled
End Sub
