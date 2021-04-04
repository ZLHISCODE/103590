VERSION 5.00
Begin VB.Form frmQueueCallCfg 
   BorderStyle     =   0  'None
   Caption         =   "排队参数配置"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmQueueCallCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkUseQueueCall 
      Caption         =   "启用排队叫号"
      Height          =   255
      Left            =   285
      TabIndex        =   1
      Top             =   165
      Width           =   1410
   End
   Begin VB.Frame framQueueCall 
      Height          =   5895
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "呼叫设置(&C)"
         Height          =   375
         Left            =   5850
         TabIndex        =   20
         Top             =   5085
         Width           =   1290
      End
      Begin VB.Frame Frame3 
         Caption         =   "基本设置"
         Height          =   1710
         Left            =   420
         TabIndex        =   16
         Top             =   405
         Width           =   6705
         Begin VB.TextBox txtClearDays 
            Height          =   300
            Left            =   930
            MaxLength       =   1
            TabIndex        =   19
            Text            =   "1"
            Top             =   690
            Width           =   465
         End
         Begin VB.CheckBox chkKeepNum 
            Caption         =   "执行间改变后排队号保持不变"
            Height          =   285
            Left            =   150
            TabIndex        =   17
            Top             =   315
            Width           =   2745
         End
         Begin VB.Label Label1 
            Caption         =   "自动清除      天前的排队数据"
            Height          =   240
            Left            =   165
            TabIndex        =   18
            Top             =   735
            Width           =   2580
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "号码生成规则："
         Height          =   1560
         Left            =   420
         TabIndex        =   2
         Top             =   2370
         Width           =   6720
         Begin VB.TextBox txtLen 
            Height          =   300
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   15
            Text            =   "3"
            Top             =   600
            Width           =   465
         End
         Begin VB.OptionButton optNum 
            Caption         =   "顺序号"
            Height          =   285
            Left            =   3990
            TabIndex        =   11
            Top             =   600
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "日"
            Height          =   210
            Left            =   3210
            TabIndex        =   8
            Top             =   630
            Width           =   510
         End
         Begin VB.CheckBox chkMonth 
            Caption         =   "月"
            Height          =   210
            Left            =   2430
            TabIndex        =   7
            Top             =   630
            Width           =   495
         End
         Begin VB.CheckBox chkYear 
            Caption         =   "年份"
            Height          =   210
            Left            =   1470
            TabIndex        =   5
            Top             =   645
            Width           =   690
         End
         Begin VB.CheckBox chkPre 
            Caption         =   "编码前缀"
            Height          =   210
            Left            =   225
            TabIndex        =   3
            Top             =   645
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.Label Label11 
            Caption         =   "序号长度："
            Height          =   225
            Left            =   4905
            TabIndex        =   14
            Top             =   645
            Width           =   975
         End
         Begin VB.Label labPreview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X121205123"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   1530
            TabIndex        =   13
            Top             =   1020
            Width           =   2040
         End
         Begin VB.Label Label9 
            Caption         =   "样式预览："
            Height          =   270
            Left            =   675
            TabIndex        =   12
            Top             =   1095
            Width           =   900
         End
         Begin VB.Label Label6 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3705
            TabIndex        =   10
            Top             =   585
            Width           =   225
         End
         Begin VB.Label Label5 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2955
            TabIndex        =   9
            Top             =   585
            Width           =   225
         End
         Begin VB.Label Label4 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2145
            TabIndex        =   6
            Top             =   585
            Width           =   225
         End
         Begin VB.Label Label3 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1275
            TabIndex        =   4
            Top             =   600
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmQueueCallCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptID As Long
Private mblnRefreshed As Boolean

Public Sub zlRefresh(lngDeptID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strRule As String
             
    On Error GoTo err
    
    mlngDeptID = lngDeptID
    
    
    strSql = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "启用排队叫号"
                chkUseQueueCall.value = Val(Nvl(rsTemp!参数值))
                
                Call chkUseQueueCall_Click
            Case "排队号保持不变"
                chkKeepNum.value = Val(Nvl(rsTemp!参数值))
            Case "自动清除排队天数"
                txtClearDays.Text = Val(Nvl(rsTemp!参数值))
            Case "排队号码规则"
                strRule = Nvl(rsTemp!参数值)
                
                
                '号码规则对应如下：
                '前缀+年份+月份+日+序号+序号位数
                chkPre.value = IIf(Val(Mid(strRule, 1, 1)) <> 0, 1, 0)
                chkYear.value = IIf(Val(Mid(strRule, 2, 1)) <> 0, 1, 0)
                chkMonth.value = IIf(Val(Mid(strRule, 3, 1)) <> 0, 1, 0)
                chkDay.value = IIf(Val(Mid(strRule, 4, 1)) <> 0, 1, 0)
                

                optNum.value = True

                txtLen.Text = Val(Mid(strRule, 6, 1))
        End Select
        rsTemp.MoveNext
    Wend
    
    Call RefreshNumPreview
    
    mblnRefreshed = True
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub RefreshNumPreview()
    labPreview.Caption = ""
    
    If chkPre.value <> 0 Then labPreview.Caption = "X"
    If chkYear.value <> 0 Then labPreview.Caption = labPreview.Caption & Format(Now, "yy")
    If chkMonth.value <> 0 Then labPreview.Caption = labPreview.Caption & Format(Now, "mm")
    If chkDay.value <> 0 Then labPreview.Caption = labPreview.Caption & Format(Now, "dd")
    
    labPreview.Caption = labPreview.Caption & Lpad("1", Val(txtLen.Text), "0")
End Sub

Private Function GetNumRule() As String
    GetNumRule = IIf(chkPre.value <> 0, "1", "0") & _
                IIf(chkYear.value <> 0, "1", "0") & _
                IIf(chkMonth.value <> 0, "1", "0") & _
                IIf(chkDay.value <> 0, "1", "0") & _
                IIf(optNum.value, "1", "0") & _
                txtLen.Text
End Function


Public Sub zlSave()
    Dim i As Integer, strInput As String
    Dim strSql As String
    
    If Not mblnRefreshed Then Exit Sub      '没有刷新则不保存
    If mlngDeptID < 0 Then Exit Sub
    
      
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '启用排队叫号','" & chkUseQueueCall.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '排队号保持不变','" & chkKeepNum.value & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '自动清除排队天数','" & txtClearDays.Text & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    strSql = "ZL_影像流程参数_UPDATE( " & mlngDeptID & ", '排队号码规则','" & GetNumRule & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    

End Sub

Private Sub chkDay_Click()
    mblnRefreshed = True
    
    Call RefreshNumPreview
End Sub

Private Sub chkKeepNum_Click()
    mblnRefreshed = True
End Sub

Private Sub chkMonth_Click()
    mblnRefreshed = True
    
    Call RefreshNumPreview
End Sub

Private Sub chkPre_Click()
    mblnRefreshed = True
    
    Call RefreshNumPreview
End Sub

Private Sub chkUseQueueCall_Click()
    mblnRefreshed = True
    framQueueCall.Enabled = IIf(chkUseQueueCall.value <> 0, True, False)
    Frame2.Enabled = framQueueCall.Enabled
    Frame3.Enabled = framQueueCall.Enabled
    
    chkKeepNum.Enabled = framQueueCall.Enabled
    Label1.Enabled = framQueueCall.Enabled
    txtClearDays.Enabled = framQueueCall.Enabled
    
    chkPre.Enabled = framQueueCall.Enabled
    chkYear.Enabled = framQueueCall.Enabled
    chkMonth.Enabled = framQueueCall.Enabled
    chkDay.Enabled = framQueueCall.Enabled
    optNum.Enabled = framQueueCall.Enabled
    txtLen.Enabled = framQueueCall.Enabled
    
    Label3.Enabled = framQueueCall.Enabled
    Label4.Enabled = framQueueCall.Enabled
    Label5.Enabled = framQueueCall.Enabled
    Label6.Enabled = framQueueCall.Enabled
    Label11.Enabled = framQueueCall.Enabled
    Label9.Enabled = framQueueCall.Enabled
    
End Sub

Private Sub chkYear_Click()
    mblnRefreshed = True
    
    Call RefreshNumPreview
End Sub

Private Sub Form_Resize()
On Error GoTo ErrHandle

    framQueueCall.Left = Fix((Me.ScaleWidth - framQueueCall.Width) / 2)
Exit Sub
ErrHandle:
End Sub

Private Sub optNum_Click()
    mblnRefreshed = True
End Sub

Private Sub txtClearDays_Change()
    mblnRefreshed = True
End Sub

Private Sub txtLen_Change()
    mblnRefreshed = True
        
    Call RefreshNumPreview
End Sub
