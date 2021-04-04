VERSION 5.00
Begin VB.Form frmXWSetParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PACS参数设置"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXWSetParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtStaticImage 
      Height          =   375
      Left            =   1680
      TabIndex        =   39
      Text            =   "http://127.0.0.1:8080/KeyImage.aspx?colid0=22&colvalue0=[@STU_NO]"
      Top             =   4440
      Width           =   9120
   End
   Begin VB.ComboBox cbo3DViewType 
      Height          =   360
      ItemData        =   "frmXWSetParams.frx":038A
      Left            =   6360
      List            =   "frmXWSetParams.frx":0394
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   5640
      Width           =   1812
   End
   Begin VB.TextBox txtWebServerPath 
      Height          =   375
      Left            =   1680
      TabIndex        =   34
      Text            =   "http://127.0.0.1:8080/TakeImage.aspx?colid0=22&colvalue0=[@STU_NO]"
      Top             =   3960
      Width           =   9120
   End
   Begin VB.TextBox txtSeriesSchemeNo 
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Text            =   "2"
      Top             =   5640
      Width           =   372
   End
   Begin VB.TextBox txtStudySchemeNo 
      Height          =   375
      Left            =   1680
      TabIndex        =   31
      Text            =   "1"
      Top             =   5640
      Width           =   372
   End
   Begin VB.TextBox txtXWOracleOwner 
      Height          =   375
      Left            =   1680
      TabIndex        =   28
      Text            =   "zlhis"
      Top             =   5040
      Width           =   2160
   End
   Begin VB.TextBox txtImageShare 
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Text            =   "DCMSHARE"
      Top             =   5040
      Width           =   1788
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "记录日志"
      Height          =   255
      Left            =   9120
      TabIndex        =   25
      Top             =   5085
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "PACS 用户设置"
      Height          =   2295
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   10575
      Begin VB.Frame Frame6 
         Caption         =   "光盘刻录"
         Height          =   1455
         Left            =   7080
         TabIndex        =   20
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtDVDBurnPswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtDVDBurnUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "密码"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "用户名"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "发送图像"
         Height          =   1455
         Left            =   3600
         TabIndex        =   15
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtSendImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtSendImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   16
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "用户名"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "密码"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   900
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "删除图像"
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtDelImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtDelImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "密码"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "用户名"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   420
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PACS 数据库服务器"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   10575
      Begin VB.CommandButton Command1 
         Caption         =   "测试(&T)"
         Height          =   400
         Left            =   9360
         TabIndex        =   36
         Top             =   460
         Width           =   1000
      End
      Begin VB.TextBox txtDBServerPswd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   7320
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtDBServerUser 
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtDBServerIP 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "密码"
         Height          =   255
         Left            =   6720
         TabIndex        =   7
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "用户名"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "服务名"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   9800
      TabIndex        =   1
      Top             =   5640
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   8520
      TabIndex        =   0
      Top             =   5640
      Width           =   1000
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "关键图像地址"
      Height          =   240
      Left            =   120
      TabIndex        =   40
      Top             =   4485
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "3D观片类型"
      Height          =   240
      Left            =   5040
      TabIndex        =   37
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "WEB观片地址"
      Height          =   240
      Left            =   240
      TabIndex        =   35
      Top             =   4000
      Width           =   1320
   End
   Begin VB.Label Label14 
      Caption         =   "序列方案号"
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   5685
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "检查方案号"
      Height          =   240
      Left            =   360
      TabIndex        =   30
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "接口包拥有者"
      Height          =   240
      Left            =   105
      TabIndex        =   29
      Top             =   5100
      Width           =   1440
   End
   Begin VB.Label Label11 
      Caption         =   "历史图像共享目录"
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   5100
      Width           =   1965
   End
End
Attribute VB_Name = "frmXWSetParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function zlShowMe(frmParent As Form) As Long
'------------------------------------------------
'功能：打开新网PACS的参数设置窗口
'返回：
'------------------------------------------------
    On Error GoTo err
    
    Call fillParams
    Me.Show 1, frmParent
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Call SaveParams
    Unload Me
End Sub

Private Sub fillParams()
'------------------------------------------------
'功能：填充新网PACS的参数
'返回：
'------------------------------------------------
    Dim i As Integer
    Dim str3DViewType As String
    
    On Error GoTo err
    
    '从中联ORACLE 模块参数中获取新网的数据库服务器IP地址，用户名和密码
    txtDBServerIP = zlDatabase.GetPara("XW数据库服务器IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerUser = zlDatabase.GetPara("XW数据库服务器用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerPswd = zlDatabase.GetPara("XW数据库服务器密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")

    txtXWOracleOwner = zlDatabase.GetPara("XWOracle拥有者", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    'txtWebServerIP = zlDatabase.GetPara("XWWEB服务器IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtWebServerPath = zlDatabase.GetPara("XWWEB观片地址", glngSys, G_LNG_XWPACSVIEW_MODULE, "http://127.0.0.1:8080/TakeImage.aspx?colid0=22&colvalue0=[@STU_NO]")
    txtStaticImage = zlDatabase.GetPara("XW关键图像地址", glngSys, G_LNG_XWPACSVIEW_MODULE, "http://127.0.0.1:8080/KeyImage.aspx?colid0=22&colvalue0=[@STU_NO]")
    
    str3DViewType = zlDatabase.GetPara("XW3D观片类型", glngSys, G_LNG_XWPACSVIEW_MODULE, "Study3D")
    For i = 0 To cbo3DViewType.ListCount - 1
        If cbo3DViewType.list(i) = str3DViewType Then
            cbo3DViewType.ListIndex = i
            Exit For
        End If
    Next
    If cbo3DViewType.ListCount > 0 Then If cbo3DViewType.ListIndex < 0 Then cbo3DViewType.ListIndex = 0
    
    txtDelImageUser = zlDatabase.GetPara("XW删除图像用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDelImagePswd = zlDatabase.GetPara("XW删除图像密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtSendImageUser = zlDatabase.GetPara("XW发送图像用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtSendImagePswd = zlDatabase.GetPara("XW发送图像密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtDVDBurnUser = zlDatabase.GetPara("XW光盘刻录用户名", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDVDBurnPswd = zlDatabase.GetPara("XW光盘刻录密码", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtImageShare = zlDatabase.GetPara("XW历史图像共享目录", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    
    chkLog.value = IIf(Val(zlDatabase.GetPara("XW记录接口日志", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1, 1, 0)
    
    txtStudySchemeNo.Text = zlDatabase.GetPara("XW检查方案号", glngSys, G_LNG_XWPACSVIEW_MODULE, "1")
    txtSeriesSchemeNo.Text = zlDatabase.GetPara("XW序列方案号", glngSys, G_LNG_XWPACSVIEW_MODULE, "2")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SaveParams()
'------------------------------------------------
'功能：保存新网PACS的参数
'返回：
'------------------------------------------------
    On Error GoTo err
    
    '将新网PACS的参数设置保存到中联ORACLE 模块参数中
    Call zlDatabase.SetPara("XW数据库服务器IP", txtDBServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW数据库服务器用户名", txtDBServerUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW数据库服务器密码", txtDBServerPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XWOracle拥有者", txtXWOracleOwner.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    'Call zlDatabase.SetPara("XWWEB服务器IP", txtWebServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XWWEB观片地址", txtWebServerPath.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW关键图像地址", txtStaticImage.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW3D观片类型", cbo3DViewType.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW删除图像用户名", txtDelImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW删除图像密码", txtDelImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW发送图像用户名", txtSendImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW发送图像密码", txtSendImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW光盘刻录用户名", txtDVDBurnUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW光盘刻录密码", txtDVDBurnPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW历史图像共享目录", txtImageShare.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW记录接口日志", chkLog.value, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW检查方案号", txtStudySchemeNo.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW序列方案号", txtSeriesSchemeNo.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandle
    Call XWTestDBConnection(txtDBServerIP.Text, txtDBServerUser.Text, txtDBServerPswd.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

