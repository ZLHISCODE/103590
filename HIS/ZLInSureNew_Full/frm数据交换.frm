VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm数据交换 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "正在进行数据交换..."
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4560
   Icon            =   "frm数据交换.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timsearch 
      Interval        =   100
      Left            =   1320
      Top             =   2160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   972
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   492
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   612
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393216
      FullWidth       =   51
      FullHeight      =   41
   End
   Begin VB.Timer TimWrite 
      Interval        =   100
      Left            =   360
      Top             =   2160
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   372
      Left            =   3360
      TabIndex        =   8
      Top             =   2160
      Width           =   972
   End
   Begin VB.PictureBox PicWrite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frm数据交换.frx":000C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   5325
   End
   Begin VB.Label lbl提示尾 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1008
      TabIndex        =   4
      Top             =   1224
      Width           =   120
   End
   Begin VB.Label lbl提示头 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1008
      TabIndex        =   2
      Top             =   456
      Width           =   120
   End
   Begin VB.Label lbl医保卡号 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1008
      TabIndex        =   3
      Top             =   840
      Width           =   96
   End
   Begin VB.Label lbl等待提示 
      AutoSize        =   -1  'True
      Caption         =   "正在等待医保结算文件..."
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   2088
   End
   Begin VB.Label lbl切换程序 
      AutoSize        =   -1  'True
      Caption         =   "请先切换到医保结算程序进行相应处理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   144
      TabIndex        =   0
      Top             =   120
      Width           =   4080
   End
End
Attribute VB_Name = "frm数据交换"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String, mstrStream As String
Private mreturn As Boolean
Private mbytType As Byte
Private mlng病人ID As Long
Private strFile As String

Private Sub cmdCancle_Click()
    If MsgBox("请注意:如果医保软件已结算,请不要使用此功能,可能会造成两边金额不平。" & vbCrLf & _
        "取消结算文件读取吗？", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbYes Then
        mstrStream = ""
        Me.Hide
    End If
End Sub
Private Sub cmdOK_Click()
    Dim strIdentify As String
    Dim strAddition As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    strIdentify = "": strAddition = ""
    If mbytType = 1 Then
        Set mdomInput = New MSXML2.DOMDocument
        mdomInput.Load ("c:\njyb\zydjxx.xml")
        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
        strIdentify = nodRowset.selectSingleNode("TBR").Text & ";"                                     '0卡号
        strIdentify = strIdentify & nodRowset.selectSingleNode("TBR").Text & ";"                    '1医保号（个人编号）
        strIdentify = strIdentify & ";"                                 '2密码
        strIdentify = strIdentify & nodRowset.selectSingleNode("XM").Text & ";"                   '3姓名
        strIdentify = strIdentify & nodRowset.selectSingleNode("XB").Text & ";"                               '4性别
        strIdentify = strIdentify & ";"                             '5出生日期
        strIdentify = strIdentify & nodRowset.selectSingleNode("SFZH").Text & ";"                                 '6身份证
        strIdentify = strIdentify & ";"                              '7.单位名称(编码)
        strAddition = "0;"                                          '8.中心代码
        strAddition = strAddition & nodRowset.selectSingleNode("XH").Text & ";"                             '9.顺序号
        strAddition = strAddition & nodRowset.selectSingleNode("XZMC").Text & ";"                        '10人员身份
        strAddition = strAddition & nodRowset.selectSingleNode("ZHYE").Text & ";"                             '11帐户余额
        strAddition = strAddition & "0;"                           '12当前状态
        strAddition = strAddition & ";"                            '13病种ID
        strAddition = strAddition & "1;"                           '14在职(1,2,3)
        strAddition = strAddition & ";"                             '15退休证号
        strAddition = strAddition & ";"                             '16年龄段
        strAddition = strAddition & ";"                             '17灰度级
        strAddition = strAddition & ";"                             '18帐户增加累计
        strAddition = strAddition & "0;"                            '19帐户支出累计
        strAddition = strAddition & "0;"                            '20上年工资总额
        strAddition = strAddition & "0"                            '21住院次数累计
    
        mlng病人ID = BuildPatiInfo(1, strIdentify & strAddition, mlng病人ID, TYPE_南京市)
        '返回格式:中间插入病人ID
        If mlng病人ID > 0 Then
            mstrStream = strIdentify & mlng病人ID & ";" & strAddition
        End If
    End If
    
    TimWrite.Enabled = False
    mreturn = True
    Call DebugTool("数据已接收，隐藏；当前mbytType=" & mbytType)
    Me.Hide
'    Call Kill("C:\NJYB\zydjxx.xml")
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
    mstrFile = gstrAviPath & "\FINDFILE.AVI"
    Call aviMove
End Sub
Private Sub aviMove()
    On Error Resume Next
    With Avi
        .Open (mstrFile)
        .AutoPlay = True
        .Play
    End With
End Sub

Private Sub Timsearch_Timer()
    Dim strTemp As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim xm As String
    Select Case mbytType
        Case 1
            strFile = "C:\NJYB\ZYDJXX.XML"
        Case 9
            strFile = "C:\NJYB\CYJSD.XML"
        Case Else
            strFile = "C:\NJYB\MZJSHZ.XML"
    End Select
    If Not FileExists(strFile) Then Exit Sub
    
    strTemp = Trim(mdl南京市.readTxtFile(strFile))
    If strTemp <> "" Then
'        Timsearch.Enabled = False
        mstrStream = strTemp
        cmdOK.Enabled = True
    Else
        lbl提示头.Caption = ""
        lbl医保卡号.Caption = ""
        lbl提示尾.Caption = ""
        cmdOK.Enabled = False
        Exit Sub
    End If
    
    If mbytType = 9 Then
        lbl提示头.Caption = "已发现结算文件"
        lbl医保卡号.Caption = ""
        lbl提示尾.Caption = ""
    Else
'        Set mdomInput = New MSXML2.DOMDocument
'        mdomInput.Load ("c:\njyb\mzjshz.xml")
'        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
        lbl提示头.Caption = "请检查医保卡号:"
'        lbl医保卡号.Caption = nodRowset.selectSingleNode("TBR").Text
'        xm = nodRowset.selectSingleNode("XM").Text
'        lbl提示尾.Caption = IIf(mbytType = 1, "姓名:" & xm, "是否正确！")
    End If
End Sub

Private Sub TimWrite_Timer()
    Static i As Long
    i = i + 20
    If i > PicWrite.ScaleWidth Then i = 1
    
    Call PicWrite.PaintPicture(PicWrite, i, 0, PicWrite.ScaleWidth - i, PicWrite.ScaleHeight, 0, 0, PicWrite.ScaleWidth - i, PicWrite.ScaleHeight)
    Call PicWrite.PaintPicture(PicWrite, 0, 0, i, PicWrite.ScaleHeight, PicWrite.ScaleWidth - i, 0, i, PicWrite.ScaleHeight)
End Sub

Public Function getFeeBalance(Optional bytType As Byte, Optional lng病人ID As Long) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    Me.Show 1
     
    If mreturn Then
        lng病人ID = mlng病人ID
        getFeeBalance = mstrStream
    End If
End Function
