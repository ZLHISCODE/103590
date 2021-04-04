VERSION 5.00
Begin VB.Form frmSet涪陵 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmSet涪陵.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd提取信息 
      Caption         =   "…"
      Height          =   285
      Left            =   3300
      TabIndex        =   5
      Top             =   630
      Width           =   285
   End
   Begin VB.TextBox txt医院编码 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   7
      Top             =   1020
      Width           =   1845
   End
   Begin VB.TextBox txt医保机构编码 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   4
      Top             =   630
      Width           =   1575
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "上传"
      Height          =   350
      Left            =   120
      TabIndex        =   16
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CheckBox chk床位 
      Caption         =   "上传床位信息"
      Height          =   210
      Left            =   2790
      TabIndex        =   12
      Top             =   1950
      Width           =   2295
   End
   Begin VB.CheckBox chk诊疗 
      Caption         =   "上传诊疗项目信息"
      Height          =   210
      Left            =   420
      TabIndex        =   11
      Top             =   1950
      Width           =   2295
   End
   Begin VB.CheckBox chk药品 
      Caption         =   "上传药品编码信息"
      Height          =   210
      Left            =   2790
      TabIndex        =   10
      Top             =   1620
      Width           =   2295
   End
   Begin VB.CheckBox chk疾病 
      Caption         =   "上传疾病编码信息"
      Height          =   210
      Left            =   420
      TabIndex        =   9
      Top             =   1620
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   1410
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3270
      TabIndex        =   15
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2010
      TabIndex        =   14
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   2310
      Width           =   5265
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1725
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lbl医院编码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医院编码(&H)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   630
      TabIndex        =   6
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lbl医保机构编码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保机构编码(&Y)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   690
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "号串口"
      Height          =   180
      Index           =   4
      Left            =   2130
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "当前串口(&D)"
      Height          =   180
      Index           =   3
      Left            =   630
      TabIndex        =   0
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSet涪陵"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng险类 As Long
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtEdit) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo errHand
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'端口号','" & txtEdit.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'医保机构编码','" & txt医保机构编码.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'医院编码','" & txt医院编码.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    gintComPort = txtEdit.Text
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTrans_Click()
    Dim rsTemp As New ADODB.Recordset, iLoop As Long, strTemp As String
    On Error GoTo errHand
'    gstr医保机构编码 = "500102"
'    gstr医院编码 = "5001020003"
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
        initType
        Call DebugTool("读取系统卡，调用函数：getybjgbm")
        mblnReturn = fl_getybjgbm(gstrOutPara)
        Call DebugTool("成功！")
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Sub
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
        txt医保机构编码.Text = gstr医保号
        txt医院编码.Text = gstr医院编码
    End If
    If chk疾病.Value = 1 Then
        Call DebugTool("准备上传保险病种")
        gstrSQL = "Select id as 编码,名称 From 保险病种"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk疾病.Caption = "上传疾病编码信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        Call DebugTool("保险病种记录数：" & rsTemp.RecordCount)
        While Not rsTemp.EOF
            initType
            Call DebugTool("函数名：wyyglxx；入参：" & gstr医保机构编码 & "," & gstr医院编码 & "," & "0" & "," & rsTemp!编码 & "," & rsTemp!名称 & "," & "")
            mblnReturn = fl_wyyglxx(gstr医保机构编码, gstr医院编码, "0", rsTemp!编码, rsTemp!名称, "", gstrOutPara)
            Call DebugTool("调用成功！")
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk疾病.Caption = "上传疾病编码信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk疾病.Value = 0
    End If
    If chk药品.Value = 1 Then
        gstrSQL = " Select a.类别 as 类别,a.id as 编码,a.名称 as 名称,b.药品来源 as 药品来源 " & _
                  " From 收费细目 a,药品目录 b " & _
                  " Where a.类别 In ('5','6','7') and a.ID=b.药品ID" & _
                  " And (A.撤档时间 Is NULL Or to_char(A.撤档时间,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk药品.Caption = "上传药品编码信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = fl_wyyglxx(gstr医保机构编码, gstr医院编码, "1", rsTemp!类别 & "_" & rsTemp!编码, rsTemp!名称, IIf(rsTemp!药品来源 = "进口", "03", "02"), gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk药品.Caption = "上传药品编码信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk药品.Value = 0
    End If
    If chk诊疗.Value = 1 Then
        gstrSQL = "select * from 收费细目 where 类别 Not In ('J','5','6','7')" & _
            " And (撤档时间 Is NULL Or to_char(撤档时间,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk诊疗.Caption = "上传诊疗项目信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = fl_wyyglxx(gstr医保机构编码, gstr医院编码, "2", rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk诊疗.Caption = "上传诊疗项目信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk诊疗.Value = 0
    End If
    If chk床位.Value = 1 Then
        gstrSQL = "select * from 收费细目 where 类别='J'" & _
            " And (撤档时间 Is NULL Or to_char(撤档时间,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk床位.Caption = "上传床位信息(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = fl_wyyglxx(gstr医保机构编码, gstr医院编码, "3", rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, " ", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk床位.Caption = "上传床位信息(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk床位.Value = 0
    End If
    MsgBox "基本项目信息上传完成", vbInformation, gstrSysName
    Exit Sub
errHand:
    MsgBox "错误号：" & Err.Number & vbCrLf & "错误信息：" & Err.Description
End Sub

Private Sub cmd提取信息_Click()
    MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
    initType
    mblnReturn = fl_getybjgbm(gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
            GoTo CheckCard
        Else
            Exit Sub
        End If
    End If
    gstr医保机构编码 = gstrOutPara.out1
    gstr医院编码 = gstrOutPara.out2
    txt医保机构编码.Text = gstr医保机构编码
    txt医院编码.Text = gstr医院编码
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnReturn = False
    
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取参数", mlng险类)
    
    With rsTemp
        Do While Not .EOF
            If !参数名 = "端口号" Then
                txtEdit.Text = Nvl(!参数值)
            ElseIf !参数名 = "医保机构编码" Then
                txt医保机构编码.Text = Nvl(!参数值)
            Else
                txt医院编码.Text = Nvl(!参数值)
            End If
            .MoveNext
        Loop
    End With
End Sub

Public Function ShowME(ByVal lng险类 As Long) As Boolean
    mlng险类 = lng险类
    Me.Show 1
    ShowME = mblnReturn
End Function
