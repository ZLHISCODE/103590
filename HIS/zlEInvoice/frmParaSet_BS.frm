VERSION 5.00
Begin VB.Form frmParaSet_BS 
   BorderStyle     =   0  'None
   Caption         =   "博思电子票据参数配置"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frmPaperCode 
      Caption         =   "纸质票据代码"
      Height          =   1005
      Left            =   90
      TabIndex        =   13
      Top             =   1920
      Width           =   6345
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   3
         Left            =   4530
         MaxLength       =   30
         TabIndex        =   38
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   2
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   16
         Top             =   600
         Width           =   1605
      End
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   0
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   14
         Top             =   270
         Width           =   1605
      End
      Begin VB.TextBox txtPaperCode 
         Height          =   285
         Index           =   1
         Left            =   4530
         MaxLength       =   100
         TabIndex        =   15
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "预交票据代码"
         Height          =   285
         Index           =   3
         Left            =   3390
         TabIndex        =   39
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "结账票据代码"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "收费票据代码"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lblPaperCode 
         Caption         =   "挂号票据代码"
         Height          =   315
         Index           =   1
         Left            =   3390
         TabIndex        =   35
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame fra误差费 
      Caption         =   "误差费控制"
      Height          =   675
      Left            =   120
      TabIndex        =   30
      Top             =   4620
      Width           =   6345
      Begin VB.TextBox txt误差费对照名称 
         Height          =   285
         Left            =   4500
         MaxLength       =   100
         TabIndex        =   34
         Top             =   270
         Width           =   1665
      End
      Begin VB.TextBox txt误差费对照编码 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   32
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label lbl误差费名称 
         Caption         =   "误差费对照名称"
         Height          =   315
         Left            =   3180
         TabIndex        =   33
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lbl误差费 
         Caption         =   "误差费对照编码"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.CheckBox chk零费用开票 
      Caption         =   "零费用开具电子票据"
      Height          =   375
      Left            =   3420
      TabIndex        =   29
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CheckBox chk录入冲红原因 
      Caption         =   "是否由操作员录入票据冲红原因"
      Height          =   375
      Left            =   150
      TabIndex        =   26
      Top             =   5280
      Width           =   3105
   End
   Begin VB.Frame fra卡类别设置 
      Caption         =   "卡类别设置"
      Height          =   1530
      Left            =   75
      TabIndex        =   17
      Top             =   2955
      Width           =   6375
      Begin VB.TextBox txt缺省卡编号 
         Height          =   285
         Left            =   5145
         TabIndex        =   28
         Text            =   "99998"
         Top             =   290
         Width           =   1050
      End
      Begin VB.TextBox txtCardNO 
         Height          =   300
         Left            =   5145
         TabIndex        =   25
         Text            =   "-"
         Top             =   1020
         Width           =   1050
      End
      Begin VB.TextBox txtNotCardCode 
         Height          =   315
         Left            =   1950
         TabIndex        =   23
         Text            =   "99999"
         Top             =   1050
         Width           =   1605
      End
      Begin VB.TextBox txtIDCardCode 
         Height          =   285
         Left            =   1950
         TabIndex        =   21
         Text            =   "99998"
         Top             =   675
         Width           =   1605
      End
      Begin VB.ComboBox cbo缺省卡类别 
         Height          =   300
         Left            =   1350
         TabIndex        =   19
         Text            =   "cbo缺省卡类别"
         Top             =   290
         Width           =   2220
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         Caption         =   "身分证号作卡类别编号"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   27
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label lblNotCardNo 
         AutoSize        =   -1  'True
         Caption         =   "无卡卡号固定传"
         Height          =   180
         Left            =   3840
         TabIndex        =   24
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         Caption         =   "病人无卡的卡类型编号"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   1110
         Width           =   1800
      End
      Begin VB.Label lblCard 
         AutoSize        =   -1  'True
         Caption         =   "缺省卡类别编号"
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   20
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label lblCard 
         Caption         =   "缺省卡类别(&D)"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   18
         Top             =   350
         Width           =   1320
      End
   End
   Begin VB.ComboBox cboContentType 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0000
      Left            =   4305
      List            =   "frmParaSet_BS.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1200
      Width           =   2070
   End
   Begin VB.ComboBox cboChar 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0011
      Left            =   1200
      List            =   "frmParaSet_BS.frx":0018
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1620
      Width           =   1605
   End
   Begin VB.TextBox txtKey 
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   825
      Width           =   5190
   End
   Begin VB.ComboBox cboVersion 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0022
      Left            =   1200
      List            =   "frmParaSet_BS.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1230
      Width           =   1635
   End
   Begin VB.TextBox txtAppID 
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   465
      Width           =   5190
   End
   Begin VB.ComboBox cboURLType 
      Height          =   300
      ItemData        =   "frmParaSet_BS.frx":0033
      Left            =   315
      List            =   "frmParaSet_BS.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   825
   End
   Begin VB.TextBox txtAddress 
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Text            =   "http://<ip>:<port>/<service>/api/medical"
      Top             =   75
      Width           =   5190
   End
   Begin VB.Label lblContentType 
      AutoSize        =   -1  'True
      Caption         =   "数据传输方式(&T)"
      Height          =   180
      Left            =   2970
      TabIndex        =   9
      Top             =   1275
      Width           =   1350
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      Caption         =   "编码字符集(&B)"
      Height          =   180
      Left            =   15
      TabIndex        =   11
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      Caption         =   "支持版本(&V)"
      Height          =   180
      Left            =   150
      TabIndex        =   7
      Top             =   1290
      Width           =   990
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      Caption         =   "签名私钥(&K)"
      Height          =   210
      Left            =   150
      TabIndex        =   5
      Top             =   870
      Width           =   990
   End
   Begin VB.Label lblAppID 
      AutoSize        =   -1  'True
      Caption         =   "应用帐号(&I)"
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   525
      Width           =   990
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      Caption         =   "UR&L"
      Height          =   180
      Left            =   30
      TabIndex        =   0
      Top             =   135
      Width           =   270
   End
End
Attribute VB_Name = "frmParaSet_BS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrAddress = "<ip>:<port>/<service>/api/medical"
Private mblnNotChange As Boolean
Private Const mstrInterfanceName = "博思电子票据平台" '接口名

Private Enum PInv_Code
    Pc_收费 = 0
    Pc_挂号
    Pc_结账
    Pc_预交
End Enum

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化参数
    '编制:刘兴洪
    '日期:2020-04-08 10:26:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strText As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    With cboURLType
        .Clear
        .AddItem "http": .ListIndex = .NewIndex
        .AddItem "https"
    End With
   
    With cboVersion
        .Clear
        .AddItem "V2.0.3":  .ListIndex = .NewIndex
        .AddItem "V3.1.0"
    End With
    
    With cboChar
        '编码
        .Clear
        .AddItem "UTF8":  .ListIndex = .NewIndex
    End With
    
    With cboContentType
        .Clear
        .AddItem "application/json":  .ListIndex = .NewIndex
    End With
    
    strSql = "Select ID,编码,名称 From 医疗卡类别 where 是否启用=1 Order by 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cbo缺省卡类别
        .Clear
        .AddItem "缺省卡医疗卡"
        .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        Do While Not rsTemp.EOF
            .AddItem rsTemp!编码 & "-" & Nvl(rsTemp!名称)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID))
            rsTemp.MoveNext
        Loop
    End With
    
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',1,'URL_Type','HTTP',NULL);
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',2,'URL_Address','','<ip>:<port>/<service>/api/medical/接口服务标识');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',3,'应用帐号','','即Appid');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',4,'签名私钥','','即KEY值');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',5,'支持版本','V2.0.3','目前只支持:V2.0.3和V3.1.0');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',6,'数据传输方式','','提交和返回数据可以为JSON格式（Content-Type: application/json）');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',7,'字符编码','UTF-8','统一采用UTF-8字符编码');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',8,'缺省卡类别ID','','缺省读取的卡类别');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',9,'医疗卡类型编号','','缺省卡类别的编号');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',10,'身份证作卡类型编号','999998','使用身份证作为上传的卡类型的编号');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',11,'病人无卡的卡类别编号','999999','病人无任何卡时上传的卡类型编号');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',12,'病人无卡的卡号','-','病人无任何卡时上传的卡号');
    'insert into 三方接口配置 ( 接口名,参数号,参数名,参数值,说明) values ('博思电子票据平台',13,'录入冲红原因','1','票据冲红时是否让弹框由操作员录入冲红原因');
    
    On Error GoTo errHandle
    
    strSql = "Select 接口名,参数号,upper(参数名) as 参数名,参数值,说明 From 三方接口配置  where 接口名='" & mstrInterfanceName & "'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            strText = Nvl(rsTemp!参数值)
            
            Select Case Nvl(!参数名)
            Case "缺省卡类别ID"
                For i = 0 To cbo缺省卡类别.ListCount
                    If cbo缺省卡类别.ItemData(i) = Val(Nvl(!参数值)) Then cbo缺省卡类别.ListIndex = i: Exit For
                Next
            Case "医疗卡类型编号"
                txt缺省卡编号.Text = Nvl(!参数值)
            Case "身份证作卡类型编号"
                txtIDCardCode.Text = Nvl(!参数值)
            Case "病人无卡的卡类别编号"
                txtNotCardCode.Text = Nvl(!参数值)
            Case "病人无卡的卡号"
                txtCardNO.Text = Nvl(!参数值)
            Case UCase("URL_Type")
                If strText <> "" Then
                    Call zlControl.CboLocate(cboURLType, strText)
                    If cboURLType.ListIndex < 0 Then cboURLType.ListIndex = 0
                End If
            Case "应用帐号"
                txtAppID.Text = strText
            Case "签名私钥"
                txtKey.Text = strText
            Case "支持版本"
                If strText <> "" Then
                    Call zlControl.CboLocate(cboVersion, strText)
                    If cboVersion.ListIndex < 0 Then cboVersion.ListIndex = 0
                End If
            Case "字符编码"
                If strText <> "" Then
                    Call zlControl.CboLocate(cboChar, strText)
                    If cboChar.ListIndex < 0 Then cboChar.ListIndex = 0
                End If
            Case "数据传输方式"
                If strText <> "" Then
                    Call zlControl.CboLocate(cboContentType, strText)
                    If cboContentType.ListIndex < 0 Then cboContentType.ListIndex = 0
                End If
            Case UCase("URL_Address")
                 mblnNotChange = True
                If strText <> "" Then
                    txtAddress.Text = strText
                    txtAddress.ForeColor = Me.ForeColor
                Else
                    txtAddress.Text = mstrAddress
                    txtAddress.ForeColor = &HC0C0C0
                End If
                 mblnNotChange = False
            Case "录入冲红原因"
                chk录入冲红原因.Value = Val(strText)
            Case "误差费对照编码"
                txt误差费对照编码.Text = strText
                txt误差费对照编码.ToolTipText = Nvl(!说明)
            Case "误差费对照名称"
                txt误差费对照名称.Text = strText
                txt误差费对照编码.ToolTipText = Nvl(!说明)
            Case "零费用开具电子票据"
                chk零费用开票.Value = Val(strText)
                chk零费用开票.ToolTipText = Nvl(!说明)
            Case "收费纸质票据代码"
                txtPaperCode(Pc_收费).Text = strText
            Case "挂号纸质票据代码"
                txtPaperCode(Pc_挂号).Text = strText
            Case "结账纸质票据代码"
                txtPaperCode(Pc_结账).Text = strText
            Case "预交纸质票据代码"
                txtPaperCode(Pc_预交).Text = strText
            End Select
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitData
End Sub

Private Sub txtAddress_Change()
    If mblnNotChange Then Exit Sub
    If txtAddress.Text = "" Or txtAddress.Text = mstrAddress Then
        txtAddress.ForeColor = &HC0C0C0
        Exit Sub
    End If
    txtAddress.ForeColor = Me.ForeColor
End Sub

Private Sub txtAddress_GotFocus()
    If Not (txtAddress.Text = "" Or txtAddress.Text = mstrAddress) Then Exit Sub
    
    mblnNotChange = True
    txtAddress.Text = ""
    txtAddress.ForeColor = Me.ForeColor
    mblnNotChange = False
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If InStr("'[]，。‘：；,.'［］", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtAddress_LostFocus()
    If Not (txtAddress.Text = "" Or txtAddress.Text = mstrAddress) Then Exit Sub
    mblnNotChange = True
    txtAddress.Text = mstrAddress
    txtAddress.ForeColor = &HC0C0C0
    mblnNotChange = False
End Sub
Private Function SaveParaValue(ByVal 参数名_In As String, ByVal str参数值 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数值
    '入参:参数名_In-可以是参数号和参数名
    '     str参数值-保存的参数值
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-04-08 12:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    '  Zl_三方接口配置_Set
    strSql = "Zl_三方接口配置_Set("
    '接口名_In 三方接口配置.接口名%Type,
    strSql = strSql & "'" & mstrInterfanceName & "',"
    '参数_In   三方接口配置.参数名%Type,
    strSql = strSql & "'" & 参数名_In & "',"
    '参数值_In 三方接口配置.参数值%Type
    strSql = strSql & "'" & str参数值 & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    SaveParaValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlSavePara() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数设置
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-04-08 12:04:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call SaveParaValue(UCase("URL_Type"), cboURLType.Text)
    Call SaveParaValue(UCase("URL_Address"), IIf(Trim(txtAddress.Text) = mstrAddress, "", Trim(txtAddress.Text)))
    Call SaveParaValue(UCase("应用帐号"), Trim(txtAppID.Text))
    Call SaveParaValue(UCase("签名私钥"), Trim(txtKey.Text))
    Call SaveParaValue(UCase("支持版本"), Trim(cboVersion.Text))
    Call SaveParaValue(UCase("数据传输方式"), Trim(cboContentType.Text))
    Call SaveParaValue(UCase("字符编码"), Trim(cboChar.Text))
    Call SaveParaValue(UCase("缺省卡类别ID"), cbo缺省卡类别.ItemData(cbo缺省卡类别.ListIndex))
    Call SaveParaValue(UCase("医疗卡类型编号"), Trim(txt缺省卡编号.Text))
    Call SaveParaValue(UCase("身份证作卡类型编号"), Trim(txtIDCardCode.Text))
    Call SaveParaValue(UCase("病人无卡的卡类别编号"), Trim(txtNotCardCode.Text))
    Call SaveParaValue(UCase("病人无卡的卡号"), Trim(txtCardNO.Text))
    Call SaveParaValue(UCase("录入冲红原因"), chk录入冲红原因.Value)
    Call SaveParaValue("误差费对照编码", Trim(txt误差费对照编码.Text))
    Call SaveParaValue("误差费对照名称", Trim(txt误差费对照名称.Text))
    Call SaveParaValue("零费用开具电子票据", chk零费用开票.Value)
    Call SaveParaValue("收费纸质票据代码", txtPaperCode(Pc_收费).Text)
    Call SaveParaValue("挂号纸质票据代码", txtPaperCode(Pc_挂号).Text)
    Call SaveParaValue("结账纸质票据代码", txtPaperCode(Pc_结账).Text)
    Call SaveParaValue("预交纸质票据代码", txtPaperCode(Pc_预交).Text)
    zlSavePara = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtAppID_KeyPress(KeyAscii As Integer)
    If InStr("'[]，。‘：；,.'［］", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
    If InStr("'[]，。‘：；,.'［］", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
