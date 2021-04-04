VERSION 5.00
Begin VB.Form frm请求等待_神木大兴 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请等待……"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "frmWait"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1935
      Top             =   -90
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frm请求等待_神木大兴.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   0
      Top             =   945
      Width           =   5025
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3435
      Top             =   -150
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门 诊"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   555
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frm请求等待_神木大兴.frx":096C
      Stretch         =   -1  'True
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl内容 
      BackStyle       =   0  'Transparent
      Caption         =   "已经提交请求，正在等待中心响应...."
      Height          =   180
      Left            =   1020
      TabIndex        =   1
      Top             =   450
      Width           =   4140
   End
   Begin VB.Label lblBack 
      BackColor       =   &H8000000A&
      Height          =   630
      Left            =   -30
      TabIndex        =   2
      Top             =   1035
      Width           =   5895
   End
End
Attribute VB_Name = "frm请求等待_神木大兴"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType  As Byte   '0-门诊,1-住院
Private mstr卡号 As String   'IC卡号
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Load()
    lblInfor.Caption = Decode(mbytType, 0, "门 诊", "住 院")
End Sub

Private Sub Timer1_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub
Public Function ShowWait(ByVal bytType As Byte, ByVal str卡号 As String) As Boolean
    '功能:显示等待窗体
    'bytType :0-门诊,1-住院
    mbytType = bytType
    mstr卡号 = str卡号
    Me.Show 1
    ShowWait = mblnOK
End Function

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    If ISCHECKDATA = True Then
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    
    Timer2.Enabled = True
End Sub
Private Function ISCHECKDATA() As Boolean
    '功能:检查结算数据
    Dim rsTemp As New ADODB.Recordset
    DebugTool "开始检查结算信息"
    
    ISCHECKDATA = False
    Select Case mbytType
    Case 0  '门诊
        gstrSQL = "" & _
            "   Select  ybkh 医保卡号, cfbh 处方编号, jssj 结算时间, jsbz 医保结算标志, " & _
            "           fyhj 本次总费用, kszf 卡上支付, tczf 统筹支付, ybje 应补现金额, xm 病人姓名   " & _
            "   From MZ_JSLSB  " & _
            "   Where upper(JSBZ) ='T' and ybkh='" & mstr卡号 & "'"
    Case Else   '住院
        gstrSQL = "" & _
           "   Select  ybkh 医保卡号, zybh 住院编号, rysj 入院时间, cysj 结算时间, jsbz 医保结算标志, tpbz 医保退票标志, " & _
           "           yybz 医院结算标志, fyhj 本次总费用, kszf 卡上支付, tczf 统筹支付, gwycb 公务员床补," & _
           "           yj 押金总额, ybje 应补现金额, gfcwf 公费床位费, zfcwf 自费床位费, gftwf 公费调温费, zftwf 自费调温费 " & _
           "   from zy_jslsb   " & _
           "   where upper(JSBZ)='T'  and ybkh='" & mstr卡号 & "'"
    End Select
    OpenRecordset_神木大兴 rsTemp, "获取结算信息", gstrSQL
    If rsTemp.EOF Then
        Exit Function
    End If
    DebugTool "检查结算信息成功"
    ISCHECKDATA = True
End Function
