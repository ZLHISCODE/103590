VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm医保结算对帐_内江 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保对帐"
   ClientHeight    =   4395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7470
   Icon            =   "frm医保结算对帐.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTCDQBM 
      Height          =   300
      ItemData        =   "frm医保结算对帐.frx":000C
      Left            =   5190
      List            =   "frm医保结算对帐.frx":000E
      TabIndex        =   23
      Top             =   225
      Width           =   1830
   End
   Begin VB.ComboBox cmbDZLB 
      Height          =   300
      ItemData        =   "frm医保结算对帐.frx":0010
      Left            =   1680
      List            =   "frm医保结算对帐.frx":001A
      TabIndex        =   22
      Text            =   "门诊"
      Top             =   1275
      Width           =   1830
   End
   Begin VB.CommandButton cmd对帐 
      Caption         =   "对帐"
      Height          =   375
      Left            =   2155
      TabIndex        =   21
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   135
      TabIndex        =   14
      Top             =   2280
      Width           =   7095
      Begin VB.TextBox txtDZJE 
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
         Left            =   5070
         TabIndex        =   19
         Top             =   795
         Width           =   1830
      End
      Begin VB.TextBox txtDZCOUNT 
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
         Left            =   1560
         TabIndex        =   17
         Top             =   795
         Width           =   1830
      End
      Begin VB.TextBox txtDZQK 
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
         Left            =   1560
         TabIndex        =   15
         Top             =   270
         Width           =   5310
      End
      Begin VB.Label Label10 
         Caption         =   "对帐金额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3675
         TabIndex        =   20
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "对帐数据条数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   18
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "对帐情况"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   300
         Width           =   1365
      End
   End
   Begin VB.TextBox txtJE 
      Enabled         =   0   'False
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
      Left            =   5190
      TabIndex        =   12
      Top             =   1800
      Width           =   1830
   End
   Begin VB.TextBox txtCOUNT 
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   10
      Top             =   1800
      Width           =   1830
   End
   Begin MSComCtl2.DTPicker dtpKSRQ 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   750
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   78643200
      CurrentDate     =   38646
   End
   Begin VB.TextBox txtHOSPID 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   225
      Width           =   1830
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   5805
      TabIndex        =   1
      Top             =   3855
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "生成"
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   3855
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpZZRQ 
      Height          =   300
      Left            =   5190
      TabIndex        =   8
      Top             =   750
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   78643200
      CurrentDate     =   38646
   End
   Begin VB.Label Label7 
      Caption         =   "上传总额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   13
      Top             =   1845
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "上传数据条数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   11
      Top             =   1845
      Width           =   1365
   End
   Begin VB.Label Label5 
      Caption         =   "终止日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   9
      Top             =   780
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "开始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   7
      Top             =   780
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "对帐类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   5
      Top             =   1305
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "统筹地区编码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   4
      Top             =   255
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "医院编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   3
      Top             =   255
      Width           =   1365
   End
End
Attribute VB_Name = "frm医保结算对帐_内江"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmd对帐_Click()
    Dim StrInput As String, strOutput As String
    Dim strArr
    '龚智毅 20051027
    If txtCOUNT.Text = "" Then Exit Sub
    If 医保初始化_成都内江 = False Then Exit Sub
    StrInput = txtHOSPID & vbTab & cmbTCDQBM & vbTab & IIf(cmbDZLB.Text = "门诊", 0, 1) & vbTab & _
               Format(dtpKSRQ, "yyyyMMdd") & vbTab & Format(dtpZZRQ, "yyyyMMdd") & vbTab & _
               Lpad(txtCOUNT * 100, 10, "0") & vbTab & Lpad(txtJE * 100, 10, "0")
    '调用对帐
    Call DebugTool("准备调用对帐")
    If 业务请求_成都内江(网上对帐_内江, StrInput, strOutput) = False Then Exit Sub
    Call DebugTool("调用对帐结束")
    strArr = Split(strOutput, vbTab)
    Select Case strArr(0)
        Case 0
            txtDZQK = "0 成功"
        Case 1
            txtDZQK = "1 金额相等,条数不等"
        Case 2
            txtDZQK = "2 金额不等,条数相等"
        Case 3
            txtDZQK = "3 金额不等,条数不等"
        Case Else
            txtDZQK = strArr(0)
    End Select
    txtDZCOUNT = strArr(1) / 100
    txtDZJE = strArr(2) / 100
    If mblnInit = False Then 医保初始化_成都内江
    
    gstrSQL = "ZL_对帐日志_INSERT('" & StrInput & "','" & strOutput & "')"
    ExecuteProcedure_ZLNJ "保存对帐日志"
End Sub

Private Sub Form_Load()
    Dim rsTC As New ADODB.Recordset
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_成都内江
    Call OpenRecordset(rsTC, "医院编码")
    txtHOSPID = Rpad(rsTC!医院编码, 5)
    
    gstrSQL = "Select Distinct substr(退休证号,1,instr(退休证号,'|')-1) As 统筹地区编码 From 保险帐户 Where 险类=" & TYPE_成都内江
    Call OpenRecordset(rsTC, "取地区码")
    Do Until rsTC.EOF
        If cmbTCDQBM.Text = "" Then
            cmbTCDQBM.Text = Nvl(rsTC!统筹地区编码)
        End If
        cmbTCDQBM.AddItem Nvl(rsTC!统筹地区编码)
        rsTC.MoveNext
    Loop
End Sub

Private Sub OKButton_Click()
    Dim rsDz As New ADODB.Recordset
    Dim rsCount As New ADODB.Recordset
    Dim curJE As Currency
    Dim lngCount As Long
    
    If cmbDZLB.Text = "门诊" Then
        gstrSQL = "Select  B.支付顺序号,sum(nvl(A.冲预交,0)) 金额 From 病人预交记录 A,保险结算记录 B,保险帐户 C,结算方式 D " & _
                  " Where A.结算方式=D.名称 And D.性质 between 3 and 4  " & _
                  " And A.结帐ID=B.记录ID And B.性质=1 And B.险类=" & TYPE_成都内江 & _
                  " And B.病人ID=C.病人ID And substr(C.退休证号,1,instr(C.退休证号,'|')-1)='" & Trim(cmbTCDQBM.Text) & "'" & _
                  " And A.收款时间 between to_date('" & Format(dtpKSRQ, "yyyy-MM-dd") & "','YYYY-MM-DD') And to_date('" & _
                  Format(dtpZZRQ + 1, "yyyy-MM-dd") & "','YYYY-MM_DD')" & _
                  " Group by B.支付顺序号 having sum(nvl(A.冲预交,0))<>0"
        '龚智毅 20051027
        Call OpenRecordset(rsDz, "保险结算记录")
        Do Until rsDz.EOF
            curJE = curJE + Nvl(rsDz!金额, 0)
            gstrSQL = "Select count(distinct 病人ID) as 上传记录数 From 医保消费信息 Where 医保流水号='" & rsDz!支付顺序号 & "'"
            Call OpenRecordset(rsCount, "上传记录")
            If rsCount.EOF = False Then
                lngCount = lngCount + rsCount!上传记录数
            End If
            rsDz.MoveNext
        Loop
    Else
        gstrSQL = "Select  b.记录id,sum(nvl(A.冲预交,0)) 金额 From 病人预交记录 A,保险结算记录 B,保险帐户 C,结算方式 D " & _
                  " Where A.结算方式=D.名称 And D.名称<>'生育盈亏' And D.性质 between 3 and 4  " & _
                  " And A.结帐ID=B.记录ID And B.性质=2 And B.险类=" & TYPE_成都内江 & _
                  " And B.病人ID=C.病人ID And substr(C.退休证号,1,instr(C.退休证号,'|')-1)='" & Trim(cmbTCDQBM.Text) & "'" & _
                  " And A.收款时间 between to_date('" & Format(dtpKSRQ, "yyyy-MM-dd") & "','YYYY-MM-DD') And to_date('" & _
                  Format(dtpZZRQ + 1, "yyyy-MM-dd") & "','YYYY-MM_DD')" & _
                  "group by b.记录id having sum(nvl(A.冲预交,0))<>0"
        Call OpenRecordset(rsDz, "保险结算记录")
        Do Until rsDz.EOF
           curJE = curJE + Nvl(rsDz!金额, 0)
           lngCount = lngCount + 1
           rsDz.MoveNext
        Loop
    End If
    
    txtJE = Format(curJE, "0.00")
    txtCOUNT = Format(lngCount, "0.00")
End Sub
