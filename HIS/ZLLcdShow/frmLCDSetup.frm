VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLCDSetup 
   Caption         =   "参数设置"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   Icon            =   "frmLCDSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   5235
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "数据配置"
      Height          =   1815
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   4695
      Begin VB.TextBox txtDelString 
         Height          =   270
         Left            =   1560
         TabIndex        =   23
         Text            =   "三部,门诊"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdCalledColor 
         Caption         =   "…"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "…"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   400
         Width           =   255
      End
      Begin VB.TextBox txtRect 
         Height          =   345
         Index           =   6
         Left            =   1560
         TabIndex        =   17
         Text            =   "6"
         Top             =   375
         Width           =   735
      End
      Begin VB.TextBox txtRect 
         Height          =   345
         Index           =   5
         Left            =   1560
         TabIndex        =   14
         Text            =   "2"
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "删除字符："
         Height          =   255
         Left            =   660
         TabIndex        =   22
         Top             =   1340
         Width           =   975
      End
      Begin VB.Shape shpCalled 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00408000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3480
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "已呼叫："
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   20
         Top             =   885
         Width           =   735
      End
      Begin VB.Shape shpCalling 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3480
         Top             =   400
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "呼叫中："
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   18
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "呼叫记录显示数："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   435
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "秒"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   885
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "轮询间隔时间："
         Height          =   255
         Left            =   310
         TabIndex        =   13
         Top             =   885
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   3240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "字体设置"
      Height          =   375
      Left            =   2655
      TabIndex        =   11
      Top             =   4575
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3855
      TabIndex        =   10
      Top             =   4575
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   1335
      TabIndex        =   9
      Top             =   4575
      Width           =   1100
   End
   Begin VB.Frame frmRect 
      Caption         =   "液晶屏位置（分辨率为单位）"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "高度："
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "宽度："
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "顶："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "左："
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   240
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLCDSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'液晶屏的参数，使用本机注册表保存

Public Function zlShowMe(frmParent As Form) As Boolean
    
    Me.Show 1, frmParent
    
    zlShowMe = True
End Function

Private Sub cmdCalledColor_Click()
    dlgColor.Color = shpCalled.FillColor
    dlgColor.ShowColor
    shpCalled.FillColor = dlgColor.Color
End Sub

Private Sub cmdCancel_Click()
    '关闭窗口
    Unload Me
End Sub

Private Sub cmdColor_Click()
    dlgColor.Color = shpCalling.FillColor
    dlgColor.ShowColor
    shpCalling.FillColor = dlgColor.Color
End Sub

Private Sub cmdFont_Click()
    Dim strReg As String
    
    On Error GoTo err
    
    strReg = "公共模块\排队叫号\液晶电视"
    dlgFont.Flags = cdlCFBoth
    dlgFont.CancelError = False  '把点取消当作错误处理
    dlgFont.FontName = GetSetting("ZLSOFT", strReg, "字体", "宋体")
    dlgFont.FontBold = GetSetting("ZLSOFT", strReg, "粗体", "False")
    dlgFont.FontItalic = GetSetting("ZLSOFT", strReg, "斜体", "False")
    dlgFont.FontSize = GetSetting("ZLSOFT", strReg, "字号", "14")
    dlgFont.ShowFont
    On Error GoTo 0
    '设置字体
    SaveSetting "ZLSOFT", strReg, "字体", dlgFont.FontName
    SaveSetting "ZLSOFT", strReg, "粗体", dlgFont.FontBold
    SaveSetting "ZLSOFT", strReg, "斜体", dlgFont.FontItalic
    SaveSetting "ZLSOFT", strReg, "字号", dlgFont.FontSize

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    '检测并保存参数
    Dim strReg As String
    
    strReg = "公共模块\排队叫号\液晶电视"
    
    SaveSetting "ZLSOFT", strReg, "左", Val(txtRect(1).Text)
    SaveSetting "ZLSOFT", strReg, "顶", Val(txtRect(2).Text)
    SaveSetting "ZLSOFT", strReg, "宽度", Val(txtRect(3).Text)
    SaveSetting "ZLSOFT", strReg, "高度", Val(txtRect(4).Text)
    SaveSetting "ZLSOFT", strReg, "LED轮询时间", Val(txtRect(5).Text)
    SaveSetting "ZLSOFT", strReg, "呼叫记录显示数", Val(txtRect(6).Text)
    SaveSetting "ZLSOFT", strReg, "呼叫中颜色", shpCalling.FillColor
    SaveSetting "ZLSOFT", strReg, "已呼叫颜色", shpCalled.FillColor
    SaveSetting "ZLSOFT", strReg, "删除字符", txtDelString.Text
    
    '关闭窗口
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    strReg = "公共模块\排队叫号\液晶电视"
    
    txtRect(1).Text = GetSetting("ZLSOFT", strReg, "左", "1024")
    txtRect(2).Text = GetSetting("ZLSOFT", strReg, "顶", "0")
    txtRect(3).Text = GetSetting("ZLSOFT", strReg, "宽度", "1024")
    txtRect(4).Text = GetSetting("ZLSOFT", strReg, "高度", "768")
    txtRect(5).Text = GetSetting("ZLSOFT", strReg, "LED轮询时间", "2")
    txtRect(6).Text = GetSetting("ZLSOFT", strReg, "呼叫记录显示数", "6")
    txtDelString.Text = GetSetting("ZLSOFT", strReg, "删除字符", "")
    shpCalling.FillColor = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
    shpCalled.FillColor = GetSetting("ZLSOFT", strReg, "已呼叫颜色", &H408000)
End Sub


Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
