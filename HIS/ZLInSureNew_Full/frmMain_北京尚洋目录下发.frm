VERSION 5.00
Begin VB.Form frmMain_北京尚洋目录下发 
   Caption         =   "下载医保就诊目录"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmMain_北京尚洋目录下发.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDown 
      Caption         =   "下载(&O)"
      Height          =   350
      Left            =   2298
      TabIndex        =   1
      Top             =   1620
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   3558
      TabIndex        =   0
      Top             =   1635
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "终止(&S)"
      Height          =   350
      Left            =   2298
      TabIndex        =   2
      Top             =   1635
      Width           =   1100
   End
   Begin VB.Label pbrBar 
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   2910
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.Label LabStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12.25%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   915
      Width           =   4560
   End
   Begin VB.Label labBar 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   885
      Width           =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      X1              =   150
      X2              =   4695
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   150
      X2              =   4665
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   150
      X2              =   4695
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   150
      X2              =   4665
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label LabCaption 
      AutoSize        =   -1  'True
      Caption         =   "下载医保就诊目录"
      Height          =   180
      Left            =   210
      TabIndex        =   3
      Top             =   270
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain_北京尚洋目录下发"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mintInsure              As Integer
Private mblnStop                As Boolean
Dim lngLoop                     As Integer
Dim strSQL                      As String
 
Const conFYXH = 0               '编号
Const conFYMC = 1               '费用名称
Const conFYDW = 2               '单位
Const conXETYPE = 3             '限额类型 0不限 1单价 2总价
Const conXE = 4                 '限额
Const conZFBL = 5               '住院自费比例
Const conMZZFBL = 6             '门诊自费比例
Const conPYDM = 7               '拼音代码
Const conFYDJ = 8               '费用单价

Public Property Let intInsure(ByVal vNewValue As Integer)
    mintInsure = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim strErrMsg           As String   '消息
    Dim lngCount            As Long
    Dim strRem              As String
    Dim strCodeID           As String
    Dim strCodeName         As String
    Dim rsTemp              As ADODB.Recordset
    Dim rsCenter            As ADODB.Recordset
    Dim lngID               As Long
On Error GoTo ErrH

    mblnStop = False
    cmdDown.Visible = False
    cmdCancel.Enabled = False
    cmdStop.Visible = True
    cmdStop.Enabled = False
    labSTATUS.Caption = "正在读取数据..."
    DoEvents
    If gcn尚洋 Is Nothing Then
        If Not 建立医保连接 Then Exit Sub
    ElseIf gcn尚洋.State = 0 Then
        If Not 建立医保连接 Then Exit Sub
    End If
    gstrSQL = "Select * From PARA_CAPTURE_ITEM" ' Where AREAID ='" & gstr医保机构编码 & "'"
    Set rsCenter = gcn尚洋.Execute(gstrSQL)
    If Not (rsCenter.EOF Or rsCenter.BOF) Then
        labBar.Visible = True
        Do While Not (rsCenter.EOF Or rsCenter.BOF)
            labSTATUS.Caption = Format(Round(rsCenter.Bookmark / rsCenter.RecordCount * 100, 2), "0.00") & " %"
            labBar.Width = rsCenter.Bookmark * pbrBar.Width / rsCenter.RecordCount
            DoEvents
            '检测类别编码在保险支付大类中是否存在
            gstrSQL = "Select count(1) from 保险支付大类 where 险类=[1] and 编码=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, Trim(rsCenter!ITEM_TYPE))
            If rsTemp.Fields(0) = 0 Then
                MsgBox "当前医保在【保险支付大类】中不存在编码【" & Trim(rsCenter!ITEM_TYPE) & "】" & vbCrLf & "请先到保险支付大类中定义该编码！" & "导入程序将被终止！", vbCritical, gstrSysName
                labSTATUS.Caption = "【保险支付大类】中不存在编码【" & Trim(rsCenter!ITEM_TYPE) & "】"
                cmdStop.Visible = True
                cmdDown.Visible = True
                cmdCancel.Enabled = True
                Exit Sub
            End If
            '其它信息写入备注
            strRem = Trim(rsCenter!ITEM_SPEC) & "—" & Trim(rsCenter!ITEM_FORM) & "—" & Trim(rsCenter!PRICE_UNIT) & "—" & "" & "—" & _
                     Trim(rsCenter!SUMARY_TYPE) & "—" & Trim(rsCenter!RECIPT_FLAG) & "—" & "" & "—" & "" & "—" & _
                     "" & "—" & "" & "—" & "" & "—" & ""
            
            '插入数据
            gstrSQL = "zl_保险项目_Insert(" & mintInsure & ",'" & rsCenter!ITEM_CODE & "','" & Replace(Trim(rsCenter!ITEM_NAME), "'", "''") & "','" & Replace(Trim(rsCenter!MNEMONIC), "'", "''") & "','" & Trim(rsCenter!ITEM_TYPE) & "','" & Trim(strRem) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            rsCenter.MoveNext
            DoEvents
        Loop
    Else
        labBar.Visible = False
        labSTATUS.Caption = "表【PARA_CAPTURE_ITEM】对象不存在数据！"
    End If
    cmdStop.Visible = True
    cmdDown.Visible = True
    cmdCancel.Enabled = True
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    cmdDown.Visible = True
    cmdCancel.Enabled = True
End Sub

Private Sub cmdStop_Click()
    mblnStop = True
End Sub

Private Sub Form_Load()
    labBar.Width = 0
    labSTATUS.Caption = ""
End Sub






