VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm门诊转医保 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊转医保"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frm门诊转医保.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   13
      Top             =   4620
      Width           =   8025
   End
   Begin VB.TextBox txt开单日期 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6570
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txt开单人 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6660
      TabIndex        =   8
      Top             =   4770
      Width           =   1215
   End
   Begin VB.CommandButton cmd上传 
      Caption         =   "上传(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5280
      TabIndex        =   7
      Top             =   4770
      Width           =   1215
   End
   Begin VB.TextBox txt总费用 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6570
      TabIndex        =   5
      Top             =   270
      Width           =   1455
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3510
      TabIndex        =   3
      Top             =   270
      Width           =   1455
   End
   Begin VB.TextBox txtNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   1
      Top             =   270
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   3195
      Left            =   120
      TabIndex        =   6
      Top             =   750
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm门诊转医保.frx":000C
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl开单日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开单日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5490
      TabIndex        =   11
      Top             =   4140
      Width           =   960
   End
   Begin VB.Label lbl开单人 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开单人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   4140
      Width           =   720
   End
   Begin VB.Label lbl总费用 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "总费用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5730
      TabIndex        =   4
      Top             =   330
      Width           =   720
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2970
      TabIndex        =   2
      Top             =   330
      Width           =   480
   End
   Begin VB.Label lblNO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   240
   End
End
Attribute VB_Name = "frm门诊转医保"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private Enum Col_T
    项目名称
    医保编码
    规格
    单位
    数量
    实收金额
    列数
End Enum

Public Sub ShowME(ByVal intinsure As Integer)
    mintInsure = intinsure
    Me.Show 1
End Sub

Private Sub CMD上传_CLICK()
    Dim lng病人ID As Long
    Dim str磁卡数据 As String
    On Error GoTo errHand
    
    If Val(txtNO.Tag) = 0 Then
        MsgBox "请先确认单据号！", vbInformation, gstrSysName
        Exit Sub
    End If
    lng病人ID = Val(txt姓名.Tag)
    
    str磁卡数据 = 身份标识_贵阳(0, lng病人ID)
    If str磁卡数据 = "" Then Exit Sub
    str磁卡数据 = Split(str磁卡数据, ";")(0)
    gcnOracle.Execute "Update 门诊费用记录 Set 病人ID=" & lng病人ID & " Where 结帐ID=" & Val(txtNO.Tag)
    If 门诊挂号_贵阳(Val(txtNO.Tag), lng病人ID) Then
        gstrSQL = " zl_门诊转医保_Insert(" & lng病人ID & "," & Val(txtNO.Tag) & ",'" & str磁卡数据 & "','" & gstrUserName & "','" & txtNO.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "门诊转医保")
        
        Call InitMsh
        txtNO.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd退出_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitMsh
End Sub

Private Sub InitMsh()
    With mshDetail
        .Clear
        .Rows = 2
        .Cols = 列数
        
        .TextMatrix(0, 项目名称) = "项目名称"
        .TextMatrix(0, 医保编码) = "医保编码"
        .TextMatrix(0, 规格) = "规格"
        .TextMatrix(0, 单位) = "单位"
        .TextMatrix(0, 数量) = "数量"
        .TextMatrix(0, 实收金额) = "实收金额"
        
        .ColWidth(项目名称) = 2000
        .ColWidth(医保编码) = 1500
        .ColWidth(规格) = 1200
        .ColWidth(单位) = 800
        .ColWidth(数量) = 1000
        .ColWidth(实收金额) = 1200
    End With
    
    txt开单人.Text = ""
    txt开单日期.Text = ""
    txt姓名.Text = ""
    txt总费用.Text = ""
End Sub

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO.Text)
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim datCurr As Date
    Dim dblMoney As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    If KeyCode <> vbKeyReturn Then Exit Sub
    StrInput = Trim(txtNO.Text)
    If Len(StrInput) < 4 Then
        datCurr = zlDatabase.Currentdate()
        txtNO.Text = PreFixNO & Format(CDate(Format(datCurr, "YYYY-MM-dd")) - CDate(Format(datCurr, "YYYY") & "-01-01") + 1, "000") & Format(StrInput, "0000") '按天顺序编号
    Else
        txtNO.Text = GetFullNO(StrInput)
    End If
    
    '提取单据明细
    gstrSQL = " Select A.结帐ID,A.病人ID,A.姓名,A.开单人,A.登记时间,B.名称 AS 项目名称,C.项目编码,B.规格,B.计算单位,A.付数*A.数次 AS 数量,A.实收金额" & _
              " From 门诊费用记录 A,收费细目 B,保险支付项目 C " & _
              " Where A.收费细目ID=B.ID And B.ID=C.收费细目ID(+) And C.险类(+)=[1]" & _
              " And Mod(A.记录性质,10)=1 And Nvl(A.实收金额,0)<>0 And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0 And A.NO=[2]" & _
              " Order by B.名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取单据明细", mintInsure, CStr(txtNO.Text))
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有找到该单据，请重新输入！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '检查病人预交记录，如果存在医保支付方式则不允许再次进行医保结算
    If ISInsure(rsTemp!结帐ID) Then
        MsgBox "已上传过医保，不允许再次进行医保结算！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsTemp
        txt姓名.Text = !姓名
        txt开单人.Text = Nvl(!开单人)
        txt开单日期.Text = Format(!登记时间, "yyyy-MM-dd")
        txtNO.Tag = !结帐ID
        txt姓名.Tag = Nvl(!病人ID, 0)
        
        Do While Not .EOF
            mshDetail.TextMatrix(.AbsolutePosition, 项目名称) = !项目名称
            mshDetail.TextMatrix(.AbsolutePosition, 医保编码) = Nvl(!项目编码)
            mshDetail.TextMatrix(.AbsolutePosition, 规格) = Nvl(!规格)
            mshDetail.TextMatrix(.AbsolutePosition, 单位) = Nvl(!计算单位)
            mshDetail.TextMatrix(.AbsolutePosition, 数量) = Format(!数量, "#0.00")
            mshDetail.TextMatrix(.AbsolutePosition, 实收金额) = Format(!实收金额, "#0.00")
            
            dblMoney = dblMoney + !实收金额
            mshDetail.Rows = mshDetail.Rows + 1
            .MoveNext
        Loop
        mshDetail.Rows = mshDetail.Rows - 1
        txt总费用.Text = Format(dblMoney, "#0.00")
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function ISInsure(ByVal lng结帐ID As Long) As String
    '如果存在医保的结算方式说明已进行过医保结算
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " select 结算方式,Nvl(冲预交,0) AS 金额 " & _
              " from 病人预交记录 A,结算方式 B " & _
              " where A.结帐ID = [1] and A.记录性质=3 and A.记录状态=1 " & _
              " And A.结算方式=B.名称 And B.性质 IN (3,4)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取HIS的预交记录结果", lng结帐ID)
    ISInsure = (rsTemp.RecordCount <> 0)
End Function
