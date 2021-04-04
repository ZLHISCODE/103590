VERSION 5.00
Begin VB.Form frmIdentify重庆 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify重庆.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton opt类别 
      BackColor       =   &H8000000A&
      Caption         =   "白内障摘除术"
      Height          =   240
      Index           =   3
      Left            =   4110
      TabIndex        =   5
      Top             =   2670
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   0
      TabIndex        =   10
      Top             =   1350
      Width           =   6660
   End
   Begin VB.Frame Frame2 
      Height          =   1785
      Left            =   3570
      TabIndex        =   11
      Top             =   1260
      Width           =   30
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -210
      TabIndex        =   9
      Top             =   2985
      Width           =   6660
   End
   Begin VB.OptionButton opt类别 
      BackColor       =   &H8000000A&
      Caption         =   "急诊抢救"
      Height          =   240
      Index           =   2
      Left            =   4110
      TabIndex        =   4
      Top             =   2310
      Width           =   1275
   End
   Begin VB.OptionButton opt类别 
      BackColor       =   &H8000000A&
      Caption         =   "特殊病门诊"
      Height          =   240
      Index           =   1
      Left            =   4110
      TabIndex        =   3
      Top             =   1950
      Width           =   1515
   End
   Begin VB.OptionButton opt类别 
      Caption         =   "普通门诊"
      Height          =   240
      Index           =   0
      Left            =   4110
      TabIndex        =   2
      Top             =   1590
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtEdit 
      Height          =   420
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1860
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   2355
      TabIndex        =   6
      Top             =   3210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   3990
      TabIndex        =   7
      Top             =   3210
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   270
      Picture         =   "frmIdentify重庆.frx":030A
      Stretch         =   -1  'True
      Top             =   210
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "重庆市医疗保险"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   2190
      TabIndex        =   8
      Top             =   495
      Width           =   3465
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   1950
      Width           =   1020
   End
End
Attribute VB_Name = "frmIdentify重庆"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr个人编号 As String
Private mint类别  As Long   '如果传入是表示0-门诊，1-住院；返回时表示11-普通门诊，13-特殊病门诊，14-急诊抢救，15-白内障摘除术，22-普通住院
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            Exit Sub
        End If
    Next
    
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "未输入个人帐户,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstr个人编号 = UCase(Trim(txtEdit(0).Text))
    If mint类别 = 0 Then
        '门诊
        If opt类别(1).Value = True Then
            mint类别 = 13
        ElseIf opt类别(2).Value = True Then
            mint类别 = 14
        ElseIf opt类别(3).Value = True Then
            mint类别 = 15
        Else
            mint类别 = 11
        End If
    Else
        '住院
        mint类别 = 21
    End If
    
    '如果是门诊则更新注册表，当前验证的病人的医保号做为下次身份验证的缺省医保号
    If mint类别 <> 21 Then
        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "上次医保号", mstr个人编号 & "|" & Format(zlDatabase.Currentdate, "yyyyMMdd"))
    End If
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetIdentify(str个人编号 As String, int类别 As Integer) As Boolean
    mblnOK = False
    mstr个人编号 = str个人编号
    mint类别 = int类别
    
    If int类别 <> 0 Then
        '非门诊登记
        opt类别(0).Enabled = False
        opt类别(1).Enabled = False
        opt类别(2).Enabled = False
        opt类别(3).Enabled = False
    End If
    frmIdentify重庆.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str个人编号 = mstr个人编号
        int类别 = mint类别
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim int参数 As Integer
    Dim strData As String
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    '获取上一次门诊验证后的医保号与日期
    strData = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "上次医保号", "")
    
    int参数 = 0
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1] And 参数名='保存医保号'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否保存上次医保号", TYPE_重庆市)
    If rsTemp.RecordCount <> 0 Then
        int参数 = Nvl(rsTemp!参数值, 0)
    End If
    
    '如果是今天，则将其设置为缺省值
    If strData <> "" And int参数 = 1 Then
        If InStr(1, strData, "|") <> 0 Then
            arrData = Split(strData, "|")
            If arrData(1) = Format(zlDatabase.Currentdate, "yyyyMMdd") Then Me.txtEdit(0).Text = UCase(arrData(0))
        End If
    End If
End Sub

Private Sub opt类别_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call cmdOK_Click
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lng病人ID As Long
    Dim int业务类型 As Integer
    Dim rsTemp As New ADODB.Recordset
    
    If Index = 0 Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        lng病人ID = GetRegisted(UCase(txtEdit(0).Text))
        If lng病人ID = 0 Then Exit Sub
        
        '重新恢复上次的业务类型
        gstrSQL = "Select 业务类型 From 保险帐户 Where 险类=[1] ANd 病人ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取业务类型", TYPE_重庆市, lng病人ID)
        If rsTemp.RecordCount <> 0 Then
            int业务类型 = Nvl(rsTemp!业务类型, 11)
            If int业务类型 = 13 Then
                opt类别(1).Value = True
            ElseIf int业务类型 = 14 Then
                opt类别(2).Value = True
            ElseIf int业务类型 = 15 Then
                opt类别(3).Value = True
            Else
                opt类别(0).Value = True
            End If
        End If
    End If
End Sub

Private Function GetRegisted(ByVal str医保号 As String) As Long
    Dim strDate As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    strStart = strDate & " 00:00:00"
    strEnd = strDate & " 23:59:59"
    '如果当天内存在就诊记录(挂号或收费)，则返回病人ID，否则返回零
    gstrSQL = " Select A.病人ID From 门诊费用记录 A,保险结算记录 B " & _
              " Where A.记录性质 In (1,4) And A.结帐ID Is Not NULL" & _
              " And A.登记时间 Between to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')" & _
              " And to_date('" & strEnd & "','yyyy-MM-dd hh24:mi:ss')" & _
              " And A.结帐ID=B.记录ID And B.性质=1" & _
              " And A.病人ID+0 =(Select 病人ID From 保险帐户 Where 险类=[1] ANd 医保号=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结帐ID", TYPE_重庆市, str医保号)
    If rsTemp.RecordCount = 0 Then Exit Function
    GetRegisted = rsTemp!病人ID
End Function
