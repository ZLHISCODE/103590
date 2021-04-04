VERSION 5.00
Begin VB.Form frmParameter 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   12660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15300
   Icon            =   "frmParameter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12660
   ScaleWidth      =   15300
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3800
      Index           =   7
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   10455
      TabIndex        =   57
      Top             =   8040
      Visible         =   0   'False
      Width           =   10455
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "签名服务器配置"
         Height          =   3600
         Index           =   7
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   10095
         Begin VB.TextBox txtShanXiURL 
            Height          =   360
            Index           =   2
            Left            =   0
            TabIndex        =   64
            Top             =   2760
            Width           =   10095
         End
         Begin VB.TextBox txtShanXiURL 
            Height          =   360
            Index           =   1
            Left            =   0
            TabIndex        =   62
            Top             =   480
            Width           =   10095
         End
         Begin VB.TextBox txtShanXiURL 
            Height          =   360
            Index           =   0
            Left            =   0
            TabIndex        =   59
            Top             =   1440
            Width           =   10095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "WSDL示例:http://117.32.132.78:9000/TSAWebService/TSASNCAEncSignActPort?wsdl"
            Height          =   180
            Left            =   0
            TabIndex        =   66
            Top             =   3240
            Width           =   6750
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "时间戳服务WSDL"
            Height          =   180
            Index           =   10
            Left            =   0
            TabIndex        =   65
            Top             =   2400
            Width           =   1260
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "医院全称"
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   63
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "签名服务WSDL"
            Height          =   180
            Index           =   9
            Left            =   0
            TabIndex        =   61
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "WSDL示例:http://111.20.164.185:8771/SNCA_CertificateAuthorityPlatform/services/CertificateAuthorityServices?wsdl"
            Height          =   180
            Left            =   0
            TabIndex        =   60
            Top             =   1920
            Width           =   10080
         End
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   6
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   4605
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   4600
      Begin VB.Frame fraPara 
         BackColor       =   &H80000005&
         Caption         =   "签名算法"
         Height          =   1695
         Index           =   6
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   4215
         Begin VB.CheckBox chkTS 
            BackColor       =   &H8000000E&
            Caption         =   "启用时间戳"
            Height          =   375
            Index           =   7
            Left            =   840
            TabIndex        =   56
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "江苏版"
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   55
            Top             =   600
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "新疆版"
            Height          =   255
            Index           =   7
            Left            =   2520
            TabIndex        =   54
            Top             =   600
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   5
      Left            =   5160
      ScaleHeight     =   2175
      ScaleWidth      =   4605
      TabIndex        =   48
      Top             =   2880
      Visible         =   0   'False
      Width           =   4600
      Begin VB.Frame fraPara 
         BackColor       =   &H80000005&
         Caption         =   "签名算法"
         Height          =   1695
         Index           =   5
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   4215
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "ESE版"
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   51
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "SEH版"
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   50
            Top             =   600
            Value           =   -1  'True
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   4
      Left            =   5040
      ScaleHeight     =   2295
      ScaleWidth      =   4605
      TabIndex        =   36
      Top             =   360
      Visible         =   0   'False
      Width           =   4600
      Begin VB.Frame fraPara 
         BackColor       =   &H80000005&
         Caption         =   "签名算法"
         Height          =   1695
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   4215
         Begin VB.CheckBox chkTS 
            BackColor       =   &H80000005&
            Caption         =   "启用签章"
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   42
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "RSA"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   39
            Top             =   600
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H80000005&
            Caption         =   "SM2"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   38
            Top             =   600
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Index           =   2
      Left            =   120
      ScaleHeight     =   2895
      ScaleWidth      =   6915
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   6915
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "签名服务器配置"
         Height          =   2280
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   6735
         Begin VB.CheckBox chkTS 
            BackColor       =   &H8000000E&
            Caption         =   "启用辽宁嘉鸿"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   6495
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "服务器URL示例:http://218.25.86.214:2010/ssoworker"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   0
            Width           =   4410
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Height          =   180
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器URL"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   3
      Left            =   9960
      ScaleHeight     =   2175
      ScaleWidth      =   4605
      TabIndex        =   31
      Top             =   5640
      Visible         =   0   'False
      Width           =   4600
      Begin VB.CheckBox chkTS 
         BackColor       =   &H8000000E&
         Caption         =   "启用时间戳"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000005&
         Caption         =   "BJCA_TS_CLIENTCOMLIB.BJCATSENGINE"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   33
         ToolTipText     =   "BJCA_TS_CLIENTCOMLIB.BJCATSENGINE"
         Top             =   840
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H80000005&
         Caption         =   "BJCA_TS_CLIENTCOM.BJCATSENGINE.1"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   32
         ToolTipText     =   "BJCA_TS_CLIENTCOM.BJCATSENGINE.1"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "时间戳控件"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Index           =   0
      Left            =   9960
      ScaleHeight     =   5295
      ScaleWidth      =   4965
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4965
      Begin VB.CheckBox chkTS 
         BackColor       =   &H8000000E&
         Caption         =   "启用签章"
         Height          =   375
         Index           =   6
         Left            =   2880
         TabIndex        =   47
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtPara 
         Height          =   480
         Index           =   7
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   4680
         Width           =   4815
      End
      Begin VB.Frame fraMethod 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1920
         TabIndex        =   40
         Top             =   2160
         Width           =   2295
         Begin VB.OptionButton optMethod 
            BackColor       =   &H8000000E&
            Caption         =   "RSA"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optMethod 
            BackColor       =   &H8000000E&
            Caption         =   "SM2"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   8
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.CheckBox chkTS 
         BackColor       =   &H8000000E&
         Caption         =   "启用时间戳"
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         Caption         =   "签名服务器配置"
         Height          =   1560
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   4815
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   0
            Left            =   1350
            MaxLength       =   16
            TabIndex        =   4
            Top             =   390
            Width           =   2895
         End
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   1
            Left            =   1350
            MaxLength       =   8
            TabIndex        =   5
            Top             =   990
            Width           =   2895
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器IP"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器端口"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   900
         End
      End
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         Caption         =   "时间戳服务器配置"
         Height          =   1560
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   2640
         Width           =   4815
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   2
            Left            =   1320
            MaxLength       =   16
            TabIndex        =   9
            Top             =   390
            Width           =   2895
         End
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   3
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   10
            Top             =   990
            Width           =   2895
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器IP"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器端口"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   900
         End
      End
      Begin VB.ComboBox cboKey 
         Height          =   300
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   37
         Width           =   2055
      End
      Begin VB.CheckBox chkTS 
         BackColor       =   &H8000000E&
         Caption         =   "启用签名服务器"
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "KEY类型"
         Height          =   180
         Index           =   1
         Left            =   2040
         TabIndex        =   45
         Top             =   97
         Width           =   630
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "扩展参数"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Top             =   4320
         Width           =   720
      End
   End
   Begin VB.PictureBox picPara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   1
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4605
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   4600
      Begin VB.Frame fraPara 
         BackColor       =   &H8000000E&
         Caption         =   "签名服务器配置"
         Height          =   2280
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   4575
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   4
            Left            =   1350
            MaxLength       =   16
            TabIndex        =   22
            Top             =   390
            Width           =   2535
         End
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   5
            Left            =   1350
            MaxLength       =   8
            TabIndex        =   21
            Top             =   990
            Width           =   2535
         End
         Begin VB.CheckBox chkTS 
            BackColor       =   &H8000000E&
            Caption         =   "启用时间戳"
            Height          =   375
            Index           =   0
            Left            =   1350
            TabIndex        =   20
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器IP"
            Height          =   180
            Index           =   6
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "服务器端口"
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   900
         End
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   15300
      TabIndex        =   0
      Top             =   12045
      Width           =   15300
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   3840
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CMD_ENUM
    CMD_OK = 0
    CMD_CANCEL = 1
End Enum

Private Enum TXT_ENUM
    TXT_SIGNIP = 0
    TXT_SIGNPORT = 1
    TXT_TSIP = 2
    TXT_TSPORT = 3
    TXT_HBCASIGNIP = 4
    TXT_HBCASIGNPORT = 5
    TXT_LNCASIGNURL = 6
    TXT_OPTION = 7
End Enum

Private Sub cboKey_Click()
    gudtPara.intKeyType = IIf(cboKey.ListIndex = -1, 0, cboKey.ListIndex)
End Sub

Private Sub chkTS_Click(Index As Integer)
    Select Case Index
    Case 3  '安信签名服务器
        gudtPara.blnIsSign = chkTS(Index).Value = vbChecked
    Case 4  '辽宁三院
        gudtPara.bytSignVersion = IIf(chkTS(Index).Value = vbChecked, 1, 0)
    Case 5, 6 '北京广西启用签章
        gudtPara.blnSignPic = chkTS(Index).Value = vbChecked
    Case 7 '北京CA江苏
        gudtPara.blnISTS = chkTS(Index).Value = vbChecked
    Case Else
        gudtPara.blnISTS = chkTS(Index).Value = vbChecked
        If Index = 1 Then
            opt(0).Enabled = gudtPara.blnISTS
            opt(1).Enabled = gudtPara.blnISTS
        End If
    End Select
End Sub

Private Sub cmdPara_Click(Index As Integer)
    Dim objCA As New clsHNCA
    Dim lngID As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnOk As Boolean
    
    If Index = CMD_OK Then
        Select Case gintCA
        Case CA_湖北
             Call HUBEI_SetParaStr
        Case CA_河北CA邯郸
            Call HBCA_SetParaStr
        Case CA_江苏CA
            Call JSCA_SetParaStr
        Case CA_北京
            Call BJCA_SetParaStr
        Case CA_新疆CA
            Call XJCA_SetParaStr
        Case CA_河南CA商丘
            Call objCA.HNCA_SetParaStr
        Case CA_北京CA广西
            Call BJCAGX_SetParaStr
        Case CA_吉林安信
            Call ANXIN_SetParaStr
        Case CA_内蒙古
            Call NMG_SetParaStr
        Case CA_网证通
            Call WZT_SetParaStr
        Case CA_上海CA
            Call SHCA_SetParaStr
        Case CA_北京CA江苏
            Call BJCAJS_SetParaStr
        Case CA_陕西省
            Call ShanXi_SetPara(Trim(txtShanXiURL(0).Text), Trim(txtShanXiURL(1).Text), Trim(txtShanXiURL(2).Text))
        End Select
    End If
    
    Unload Me
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Load()
    Dim objCA As New clsHNCA
    
    On Error GoTo ErrH
    '设置缺省值
    gudtPara.strSIGNIP = ""
    gudtPara.strSignPort = ""
    gudtPara.strTSIP = ""
    gudtPara.strTSPort = ""
    gudtPara.strSignURL = ""
    gudtPara.blnISTS = False
    gudtPara.bytSignVersion = V_RSA
    fraMethod.Visible = False
    chkTS(2).Visible = False
    chkTS(3).Visible = False
    chkTS(4).Visible = False
    
    Select Case gintCA
    Case CA_湖北, CA_深圳, CA_吉林安信, CA_内蒙古, CA_网证通
        Me.Width = 5160
        Me.Height = 4620
        picPara(0).Visible = True
        chkTS(2).Visible = False
        chkTS(6).Visible = False  '签章
        lblInfo(0).Visible = False: lblInfo(1).Visible = False
        cboKey.Visible = False
        txtPara(7).Visible = False
        Select Case gintCA
        Case CA_湖北
            Me.Width = 5160
            Me.Height = 4995
            Call HUBEI_GetPara
        Case CA_内蒙古
            Me.Height = 5000
            chkTS(6).Visible = True
            Call NMG_GetPara
            chkTS(6).Value = IIf(gudtPara.blnSignPic, vbChecked, vbUnchecked)
        Case CA_吉林安信
            chkTS(2).Visible = True
            chkTS(3).Visible = True
            txtPara(7).Visible = True
            lblInfo(0).Visible = True: lblInfo(1).Visible = True: cboKey.Visible = True
            Me.Width = 5160
            Me.Height = 6420
            Call ANXIN_GetPara
            cboKey.AddItem "飞天EPASS3000GM"
            cboKey.AddItem "龙脉GM3000"
            cboKey.ListIndex = gudtPara.intKeyType
            chkTS(2).Value = IIf(gudtPara.blnISTS, vbChecked, vbUnchecked)
            chkTS(3).Value = IIf(gudtPara.blnIsSign, vbChecked, vbUnchecked)
            txtPara(7).Text = gudtPara.strOption
            txtPara(7).ToolTipText = "CA部件名称设置(多个名称用字符(&)分隔)" & vbCrLf & _
                                     "示例:SERfR01DQUlTLmRsbA==&U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="
        Case CA_网证通
            Me.Width = 10365
            Me.Height = 6400
            txtPara(7).Visible = True: chkTS(2).Visible = True
            lblInfo(0).Visible = True
            lblInfo(0).Caption = "网关证书"
            Call WZT_GetPara
            txtPara(7).Text = gudtPara.strOption
            txtPara(7).ToolTipText = "网关证书"
            chkTS(2).Value = IIf(gudtPara.blnISTS, vbChecked, vbUnchecked)
        End Select
        txtPara(0).Text = gudtPara.strSIGNIP
        txtPara(1).Text = gudtPara.strSignPort
        txtPara(2).Text = gudtPara.strTSIP
        txtPara(3).Text = gudtPara.strTSPort
        
        If gintCA = CA_湖北 Then
            fraMethod.Visible = True
            optMethod(V_RSA).Value = gudtPara.bytSignVersion = V_RSA
            optMethod(V_SM2).Value = gudtPara.bytSignVersion = V_SM2
        End If
    Case CA_河北CA邯郸
        Me.Width = 4905
        Me.Height = 3660
        picPara(1).Visible = True
        Call HBCA_GetPara
        txtPara(4).Text = gudtPara.strSIGNIP
        txtPara(5).Text = gudtPara.strSignPort
        chkTS(0).Value = IIf(gudtPara.blnISTS, Checked, Unchecked)
    Case CA_江苏CA
        Me.Width = 7080
        Me.Height = 2970
        picPara(2).Visible = True
        Select Case gintCA
        Case CA_江苏CA
            Call JSCA_GetPara
            lblNote.Caption = "参数格式:http://202.102.85.153:8080/HealthWebService.asmx?WSDL"
        End Select
        txtPara(6).Text = gudtPara.strSignURL
    Case CA_新疆CA
        Me.Width = 7080
        Me.Height = 2970
        picPara(2).Visible = True
        Call XJCA_GetPara
        lblNote.Caption = "参数格式:http://124.117.245.71:48080/webServices/ssoService" & vbCrLf & _
                          "|4028f6d24a2d7182014a2d83333e001a|华大"
        txtPara(6).Text = gudtPara.strSignURL
    Case CA_河南CA商丘
        Me.Width = 7080
        Me.Height = 2970
        picPara(2).Visible = True
        Call objCA.HNCA_GetPara
        lblNote.Caption = "参数值格式""服务器URL|TSIP|TSPORT|SSLPORT|时间戳(0-不启用/1-启用)" & vbCrLf & _
                        "|签名算法(0-RSA\1-SM2)如:http://218.28.16.104:7080/CAServer/servlet/" & vbCrLf & _
                        "CertChechServlet|218.28.16.104|8080|443|0|0"
        txtPara(6).Text = gudtPara.strSignURL
    Case CA_北京
        Me.Height = 2970
        Me.Width = 4770
        picPara(3).Visible = True
        opt(0).Enabled = False
        opt(1).Enabled = False
        Call BJCA_GetPara
        chkTS(1).Value = IIf(gudtPara.blnISTS, Checked, Unchecked)
        If opt(0).Enabled Then opt(0).Value = Val(gudtPara.strTSVersion) = 0
        If opt(1).Enabled Then opt(1).Value = Val(gudtPara.strTSVersion) = 1
    Case CA_北京CA广西
        Me.Height = 2970
        Me.Width = 4770
        picPara(4).Visible = True
        opt(2).Enabled = True
        opt(3).Enabled = True
        chkTS(5).Enabled = True
        Call BJCAGX_GetPara
        If opt(2).Enabled Then opt(2).Value = gudtPara.bytSignVersion = V_RSA
        If opt(3).Enabled Then opt(3).Value = gudtPara.bytSignVersion = V_SM2
        If chkTS(5).Enabled Then chkTS(4).Value = IIf(gudtPara.blnSignPic, vbChecked, vbUnchecked)
    Case CA_上海CA
        Me.Height = 2970
        Me.Width = 4770
        picPara(5).Visible = True
        opt(4).Enabled = True
        opt(5).Enabled = True
        Call SHCA_GetPar
        If opt(4).Enabled Then opt(4).Value = gudtPara.bytSignVersion = V_SEH
        If opt(5).Enabled Then opt(5).Value = gudtPara.bytSignVersion = V_ESE
    Case CA_北京CA江苏
        Me.Height = 2970
        Me.Width = 4770
        picPara(6).Visible = True
        opt(6).Enabled = True
        opt(7).Enabled = True
        Call BJCAJS_GetPara
        If opt(6).Enabled Then opt(6).Value = gudtPara.bytSignVersion = V_江苏
        If opt(7).Enabled Then opt(7).Value = gudtPara.bytSignVersion = V_新疆
        If chkTS(7).Enabled Then chkTS(7).Value = IIf(gudtPara.blnISTS, vbChecked, vbUnchecked)
    Case CA_陕西省
        Me.Height = 5000
        Me.Width = 10720
        picPara(7).Visible = True
        Call ShanXi_GetPara
        txtShanXiURL(0).Text = gudtPara.strSignURL
        txtShanXiURL(1).Text = gudtPara.strOption
        txtShanXiURL(2).Text = gudtPara.strTSIP
    End Select
        
    Exit Sub
ErrH:
    MsgBoxEx "参数设置加载失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Select Case gintCA
    Case CA_湖北, CA_深圳, CA_吉林安信, CA_内蒙古, CA_网证通
        picPara(0).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    Case CA_河北CA邯郸
        picPara(1).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    Case CA_江苏CA, CA_新疆CA, CA_河南CA商丘
        picPara(2).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    Case CA_北京
        picPara(3).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    Case CA_北京CA广西
        picPara(4).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
        fraPara(4).Move 120, 120, picPara(4).ScaleWidth - 240, picPara(4).ScaleHeight - 240
    Case CA_上海CA
        picPara(5).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
        fraPara(5).Move 120, 120, picPara(5).ScaleWidth - 240, picPara(5).ScaleHeight - 240
    Case CA_北京CA江苏
        picPara(6).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
        fraPara(6).Move 120, 120, picPara(6).ScaleWidth - 240, picPara(6).ScaleHeight - 240
    Case CA_陕西省
        picPara(7).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
        fraPara(7).Move 120, 120, picPara(7).ScaleWidth - 240, picPara(7).ScaleHeight - 240
    End Select
    
    cmdPara(CMD_CANCEL).Left = picBottom.Width - 1100 - 120
    cmdPara(CMD_OK).Left = cmdPara(CMD_CANCEL).Left - 1100 - 60
End Sub

Private Sub opt_Click(Index As Integer)
    Select Case Index
    Case 0, 1
        gudtPara.strTSVersion = Index
    Case 2, 3
        gudtPara.bytSignVersion = IIf(Index = 2, 0, 1)
    Case 4, 5
        gudtPara.bytSignVersion = IIf(Index = 4, 0, 1)
    Case 6, 7
        gudtPara.bytSignVersion = IIf(Index = 6, 0, 1)
    End Select
End Sub

Private Sub optMethod_Click(Index As Integer)
    gudtPara.bytSignVersion = Index
End Sub

Private Sub picPara_Resize(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        fraPara(0).Move 0, 0, 4815, 1560
        If gintCA = CA_吉林安信 Then
            chkTS(3).Move 0, 0
            chkTS(2).Move 0, 2160
            fraPara(0).Move 0, 480, 4815, 1560
            fraPara(1).Move 0, 2640, 4815, 1560
            lblInfo(0).Move 0, fraPara(1).Top + fraPara(1).Height + 240
            txtPara(7).Move 0, lblInfo(0).Top + lblInfo(0).Height + 120, 4815, 360
        ElseIf gintCA = CA_湖北 Then
            fraMethod.Move 1920, 1680
            fraPara(1).Move 0, 2160, 4815, 1560
        ElseIf gintCA = CA_网证通 Then
            fraPara(0).Move 0, 0, 4815, 1560
            fraPara(1).Move 5160, 0, 4815, 1560
            lblInfo(0).Move 0, 1800
            txtPara(7).Move 0, 2040, 9975, 3000
            chkTS(2).Move lblInfo(0).Left + lblInfo(0).Width + 600, 1700
        ElseIf gintCA = CA_内蒙古 Then
            fraPara(0).Move 0, 0, 4815, 1560
            fraPara(1).Move 0, 1680, 4815, 1560
            chkTS(6).Move 120, fraPara(1).Top + fraPara(1).Height + 120
        Else
            fraPara(1).Move 0, 1680, 4815, 1560
        End If
    End If
End Sub

Private Sub txtPara_Change(Index As Integer)
    Select Case Index
    
    Case TXT_SIGNIP, TXT_HBCASIGNIP
        gudtPara.strSIGNIP = Trim(txtPara(Index).Text)
    Case TXT_SIGNPORT, TXT_HBCASIGNPORT
        gudtPara.strSignPort = Trim(txtPara(Index).Text)
    Case TXT_TSIP
        gudtPara.strTSIP = Trim(txtPara(Index).Text)
    Case TXT_TSPORT
        gudtPara.strTSPort = Trim(txtPara(Index).Text)
    Case TXT_LNCASIGNURL
        gudtPara.strSignURL = Trim(txtPara(Index).Text)
    Case TXT_OPTION
        gudtPara.strOption = Trim(txtPara(Index).Text)
    End Select
End Sub

Private Sub txtPara_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(",8,9,10,13,", "," & KeyAscii & ",") = 0 Then
        Select Case Index
        Case TXT_SIGNIP, TXT_TSIP, TXT_HBCASIGNIP
            If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case TXT_SIGNPORT, TXT_TSPORT, TXT_HBCASIGNPORT
            If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End Select
    End If
End Sub

