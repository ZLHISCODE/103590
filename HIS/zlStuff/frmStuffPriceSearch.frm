VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffPriceSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra范围 
      Height          =   3810
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5385
      Begin VB.CheckBox Chk药品 
         Caption         =   "卫材"
         Height          =   300
         Left            =   480
         TabIndex        =   22
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Txt药品 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton Cmd药品 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4920
         TabIndex        =   20
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox chkPrice 
         Caption         =   "仅售价调价"
         Height          =   300
         Left            =   480
         TabIndex        =   9
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CheckBox chk执行日期 
         Caption         =   "执行日期"
         Height          =   300
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox chk填制日期 
         Caption         =   "填制日期"
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txt结束NO 
         Height          =   300
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txt开始No 
         Height          =   300
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Width           =   1605
      End
      Begin VB.CheckBox chkPriceAndCost 
         Caption         =   "成本价和售价一起调价"
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Top             =   3240
         Width           =   2100
      End
      Begin VB.CheckBox chkCost 
         Caption         =   "仅成本价调价"
         Height          =   300
         Left            =   2400
         TabIndex        =   3
         Top             =   2880
         Width           =   1400
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Index           =   0
         Left            =   3585
         TabIndex        =   11
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   1845
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Index           =   1
         Left            =   3585
         TabIndex        =   13
         Top             =   1845
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   196411395
         CurrentDate     =   36263
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   0
         Left            =   3345
         TabIndex        =   19
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "填制日期"
         Height          =   180
         Index           =   0
         Left            =   900
         TabIndex        =   18
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   3
         Left            =   3345
         TabIndex        =   17
         Top             =   1905
         Width           =   180
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "执行日期"
         Height          =   180
         Index           =   1
         Left            =   900
         TabIndex        =   16
         Top             =   1905
         Width           =   720
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   420
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "调价汇总号"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmStuffPriceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrResult As String

Private Type Type_Condition '过滤时设置的日期
    date填制时间开始 As Date
    date填制时间结束 As Date
    date执行时间开始 As Date
    date执行时间结束 As Date
End Type

Private mSQLCondition As Type_Condition

Private Sub chk执行日期_Click()
    If chk执行日期.Value = 1 Then
        dtp开始时间(1).Enabled = True
        dtp结束时间(1).Enabled = True
    Else
        dtp开始时间(1).Enabled = False
        dtp结束时间(1).Enabled = False
    End If
End Sub

Private Sub Chk药品_Click()
    If Chk药品.Value = 1 Then
        Txt药品.Enabled = True
        Cmd药品.Enabled = True
    Else
        Txt药品.Enabled = False
        Cmd药品.Enabled = False
    End If
End Sub

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim date填制时间开始 As Date
    Dim date填制时间结束 As Date
    Dim date执行时间开始 As Date
    Dim date执行时间结束 As Date
    
    mstrResult = ""
    If Trim(txt开始No.Text) <> "" Then
        If IsNumeric(txt开始No.Text) Then
            If Len(txt开始No.Text) < 10 Then
                MsgBox "请输入正确的调价汇总号（全数字10位）！", vbInformation, gstrSysName
                Me.txt开始No.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "请输入正确的调价汇总号（全数字10位）！", vbInformation, gstrSysName
            Me.txt开始No.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txt结束NO.Text) <> "" Then
        If CLng(Val(Trim(txt开始No.Text))) > CLng(Val(Trim(txt结束NO.Text))) Then
            MsgBox "开始调价汇总号不能小于结束调价汇总号！", vbInformation, gstrSysName
            Me.txt开始No.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(txt结束NO.Text) <> "" Then
        If IsNumeric(txt结束NO.Text) Then
            If Len(txt结束NO.Text) < 10 Then
                MsgBox "请输入正确的调价汇总号（全数字10位）！", vbInformation, gstrSysName
                Me.txt结束NO.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "请输入正确的调价汇总号（全数字10位）！", vbInformation, gstrSysName
            Me.txt结束NO.SetFocus
            Exit Sub
        End If
    End If
    
    If Chk药品.Value = 1 Then
        If Val(Txt药品.Tag) = 0 Then
            MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
            Me.Txt药品.SetFocus
            Exit Sub
        End If
    End If
    
    '调价汇总号
    If Trim(txt开始No.Text) <> "" And Trim(txt结束NO.Text <> "") Then
        mstrResult = " a.调价号 >= " & txt开始No.Text & " and a.调价号 <= " & txt结束NO.Text
    ElseIf Trim(txt开始No.Text) <> "" And Trim(txt结束NO.Text = "") Then
        mstrResult = " a.调价号 >= " & txt开始No.Text
    ElseIf Trim(txt开始No.Text) = "" And Trim(txt结束NO.Text <> "") Then
        mstrResult = " a.调价号 <= " & txt结束NO.Text
    End If
    
    '日期
    If chk填制日期.Value = 1 And chk执行日期.Value = 0 Then
        If mstrResult = "" Then
            mstrResult = " a.填制日期 between [1] and [2] "
        Else
            mstrResult = mstrResult + " and a.填制日期 between [1] and [2] "
        End If
    ElseIf chk填制日期.Value = 0 And chk执行日期.Value = 1 Then
        If mstrResult = "" Then
            mstrResult = " a.执行日期 between [3] and [4] "
        Else
            mstrResult = mstrResult + " and a.执行日期 between [3] and [4] "
        End If
    ElseIf chk填制日期.Value = 1 And chk执行日期.Value = 1 Then
        If mstrResult = "" Then
            mstrResult = " a.填制日期 between [1] and [2] and a.执行日期 between [3] and [4] "
        Else
            mstrResult = mstrResult + " and a.填制日期 between [1] and [2] and a.执行日期 between [3] and [4] "
        End If
    End If
    '填制日期
    If chk填制日期.Value = 1 Then
        mSQLCondition.date填制时间开始 = CDate(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
        mSQLCondition.date填制时间结束 = CDate(Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    End If
    
    If chk执行日期.Value = 1 Then
        mSQLCondition.date执行时间开始 = CDate(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
        mSQLCondition.date执行时间结束 = CDate(Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
    End If
    '执行日期
    
    '其他条件
    If chkPrice.Value = 1 And chkCost.Value = 0 And chkPriceAndCost.Value = 0 Then '仅调售价
        If mstrResult = "" Then
            mstrResult = " a.类型=0 "
        Else
            mstrResult = mstrResult + " and a.类型=0 "
        End If
    ElseIf chkPrice.Value = 0 And chkCost.Value = 1 And chkPriceAndCost.Value = 0 Then '仅调成本价
        If mstrResult = "" Then
            mstrResult = " a.类型=1 "
        Else
            mstrResult = mstrResult + " and a.类型=1 "
        End If
    ElseIf chkPrice.Value = 0 And chkCost.Value = 0 And chkPriceAndCost.Value = 1 Then '成本价售价一起调整
        If mstrResult = "" Then
            mstrResult = " a.类型=2 "
        Else
            mstrResult = mstrResult + " and a.类型=2 "
        End If
    ElseIf chkPrice.Value = 1 And chkCost.Value = 1 And chkPriceAndCost.Value = 0 Then '仅调售价和仅调成本价
        If mstrResult = "" Then
            mstrResult = " (a.类型=0 or a.类型=1) "
        Else
            mstrResult = mstrResult + " and (a.类型=0 or a.类型=1) "
        End If
    ElseIf chkPrice.Value = 1 And chkCost.Value = 0 And chkPriceAndCost.Value = 1 Then '仅调售价和成本价售价一起调整
        If mstrResult = "" Then
            mstrResult = " (a.类型=0 or a.类型=2) "
        Else
            mstrResult = mstrResult + " and (a.类型=0 or a.类型=2) "
        End If
    ElseIf chkPrice.Value = 0 And chkCost.Value = 1 And chkPriceAndCost.Value = 1 Then '仅调成本价和成本价售价一起调整
        If mstrResult = "" Then
            mstrResult = " (a.类型=1 or a.类型=2) "
        Else
            mstrResult = mstrResult + " and (a.类型=1 or a.类型=2) "
        End If
    End If
    
    '药品
    If Val(Txt药品.Tag) <> 0 Then
        If mstrResult = "" Then
            mstrResult = " a.调价号 In (Select 调价汇总号 From 收费价目 Where 收费细目id = " & Txt药品.Tag & GetPriceClassString("") & _
                            " union all " & _
                         " Select  调价汇总号 From 成本价调价信息 Where 药品id = " & Txt药品.Tag & ")"
        Else
            mstrResult = mstrResult & " and a.调价号 In (Select 调价汇总号 From 收费价目 Where 收费细目id =" & Txt药品.Tag & GetPriceClassString("") & _
                        " union all " & _
                         " Select  调价汇总号 From 成本价调价信息 Where 药品id = " & Txt药品.Tag & ")"
        End If
    End If
    Unload Me
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    On Error GoTo ErrHandle
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 1, 0)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt药品 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt药品.Tag = RecReturn!材料ID
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Me.dtp结束时间(0) = Sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
End Sub

Public Sub ShowMe(ByVal FrmParent As Form, ByRef strResult As String, ByRef date填制时间开始 As Date, ByRef date填制时间结束 As Date, ByRef date审核时间开始 As Date, ByRef date审核时间结束 As Date)
    Me.Show vbModal, FrmParent
        
    strResult = mstrResult
    date填制时间开始 = mSQLCondition.date填制时间开始
    date填制时间结束 = mSQLCondition.date填制时间结束
    date审核时间开始 = mSQLCondition.date执行时间开始
    date审核时间结束 = mSQLCondition.date执行时间结束
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Txt药品.Left
    sngTop = Me.Top + Txt药品.Height + Txt药品.Top '  50
    If sngTop + 4530 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 4530
    End If
    
    strKey = Trim(Txt药品.Text)
    If Mid(strKey, 1, 1) = "[" Then
        If InStr(2, strKey, "]") <> 0 Then
            strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
        Else
            strKey = Mid(strKey, 2)
        End If
    End If
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , , , strKey, sngLeft, sngTop)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt药品 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt药品.Tag = RecReturn!材料ID
    
    Txt药品.SetFocus
End Sub

