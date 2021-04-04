VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDrugSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   Icon            =   "frmDrugSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   6015
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Fra数据配置 
      Caption         =   "区域和数据配置"
      Height          =   5895
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   5535
      Begin VB.Frame 呼叫区域 
         Caption         =   "呼叫区域"
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   5295
         Begin VB.CommandButton cmdFont 
            Caption         =   "字体设置"
            Height          =   350
            Index           =   0
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdCall 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   3720
            TabIndex        =   52
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl字体 
            Caption         =   "宋体;加粗;24"
            Height          =   225
            Index           =   0
            Left            =   1320
            TabIndex        =   54
            Top             =   303
            Width           =   2120
         End
         Begin VB.Shape shpCall 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame fra选择性显示 
         Caption         =   "选择性显示"
         Height          =   4335
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   5295
         Begin VB.CommandButton cmdPreparingColor 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   3720
            TabIndex        =   37
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   3720
            TabIndex        =   36
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chk显示待配药 
            Caption         =   "显示待配药"
            Height          =   200
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CheckBox chk显示待发药 
            Caption         =   "显示待发药"
            Height          =   200
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chk显示其他区域 
            Caption         =   "显示其他区域"
            Height          =   200
            Left            =   120
            TabIndex        =   33
            Top             =   3120
            Width           =   1575
         End
         Begin VB.CheckBox chk显示窗体 
            Caption         =   "显示窗口"
            Height          =   200
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdWin 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   3720
            TabIndex        =   31
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "字体设置"
            Height          =   350
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "字体设置"
            Height          =   350
            Index           =   2
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "字体设置"
            Height          =   350
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "字体设置"
            Height          =   350
            Index           =   4
            Left            =   240
            TabIndex        =   27
            Top             =   3360
            Width           =   975
         End
         Begin VB.CommandButton cmdOther 
            Caption         =   "字体颜色"
            Height          =   350
            Left            =   3720
            TabIndex        =   26
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox txt待配药列数 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   25
            Text            =   "3"
            Top             =   2707
            Width           =   615
         End
         Begin VB.TextBox txt待配药行数 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   720
            TabIndex        =   24
            Text            =   "1"
            Top             =   2707
            Width           =   615
         End
         Begin VB.TextBox txt待发药行数 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   720
            TabIndex        =   23
            Text            =   "1"
            Top             =   1627
            Width           =   615
         End
         Begin VB.TextBox txt待发药列数 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   22
            Text            =   "3"
            Top             =   1627
            Width           =   615
         End
         Begin VB.TextBox txtContent 
            Height          =   300
            Left            =   1200
            TabIndex        =   21
            Top             =   3840
            Width           =   3975
         End
         Begin VB.Label lbl字体 
            Caption         =   "待发药"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   50
            Top             =   1278
            Width           =   2120
         End
         Begin VB.Label lbl字体 
            Caption         =   "待配药"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   49
            Top             =   2358
            Width           =   2120
         End
         Begin VB.Shape shpWin 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lbl字体 
            Caption         =   "窗口"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   48
            Top             =   558
            Width           =   2120
         End
         Begin VB.Shape shpOther 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00408000&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   3385
            Width           =   375
         End
         Begin VB.Label lbl字体 
            Caption         =   "其他内容"
            Height          =   195
            Index           =   4
            Left            =   1320
            TabIndex        =   47
            Top             =   3438
            Width           =   2120
         End
         Begin VB.Label lbl待配药行数 
            Caption         =   "行数"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lbl待配药列数 
            Caption         =   "列数"
            Height          =   195
            Left            =   1680
            TabIndex        =   45
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lbl待配药人数 
            Caption         =   "人数"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lbl待配药Sum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   43
            Top             =   2707
            Width           =   615
         End
         Begin VB.Label lbl待发药行数 
            Caption         =   "行数"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lbl待发药人数 
            Caption         =   "人数"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lbl待发药Sum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   40
            Top             =   1627
            Width           =   615
         End
         Begin VB.Label lbl待发药列数 
            Caption         =   "列数"
            Height          =   195
            Left            =   1680
            TabIndex        =   39
            Top             =   1680
            Width           =   375
         End
         Begin VB.Shape shpPreparing 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00408000&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   2305
            Width           =   375
         End
         Begin VB.Shape shpCalling 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   1225
            Width           =   375
         End
         Begin VB.Label lblContent 
            Caption         =   "显示内容"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   3893
            Width           =   735
         End
      End
      Begin VB.TextBox txt轮询时间 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4080
         TabIndex        =   19
         Text            =   "10"
         Top             =   5490
         Width           =   615
      End
      Begin VB.TextBox txt翻页时间 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Text            =   "10"
         Top             =   5490
         Width           =   615
      End
      Begin VB.Label lbl刷新时间 
         Caption         =   "LCD刷新时间"
         Height          =   195
         Left            =   2880
         TabIndex        =   58
         Top             =   5550
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "秒"
         Height          =   195
         Left            =   2040
         TabIndex        =   57
         Top             =   5550
         Width           =   255
      End
      Begin VB.Label lbl翻页时间 
         Caption         =   "翻页时间"
         Height          =   195
         Left            =   480
         TabIndex        =   56
         Top             =   5550
         Width           =   735
      End
      Begin VB.Label lbl秒 
         Caption         =   "秒"
         Height          =   195
         Left            =   4800
         TabIndex        =   55
         Top             =   5550
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   9000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   9000
      Width           =   1100
   End
   Begin VB.Frame Fra显示窗口 
      Caption         =   "显示窗口"
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   5535
      Begin VB.Frame fra叫号窗口 
         Caption         =   "叫号窗口"
         Height          =   1215
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   3255
         Begin VB.ListBox lst叫号窗口 
            Columns         =   1
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Frame Fra显示模式 
         Caption         =   "显示模式"
         Height          =   1215
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton Opt窗口 
            Caption         =   "多窗口"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Opt窗口 
            Caption         =   "单窗口"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame frmRect 
      Caption         =   "液晶屏位置（分辨率为单位）"
      Height          =   1150
      Left            =   240
      TabIndex        =   0
      Top             =   1770
      Width           =   5535
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   310
         Width           =   1695
      End
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   3
         Top             =   310
         Width           =   1695
      End
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   3
         Left            =   840
         TabIndex        =   2
         Top             =   710
         Width           =   1695
      End
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   4
         Left            =   3600
         TabIndex        =   1
         Top             =   710
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "左："
         Height          =   255
         Index           =   0
         Left            =   405
         TabIndex        =   8
         Top             =   340
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "顶："
         Height          =   255
         Index           =   1
         Left            =   3165
         TabIndex        =   7
         Top             =   345
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "宽度："
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "高度："
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   750
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   0
      Left            =   480
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   1
      Left            =   2040
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   2
      Left            =   2520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   3
      Left            =   3120
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   4
      Left            =   1440
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDrugSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrWins As String

Private Sub chk显示窗体_Click()
    If chk显示窗体.Value = 1 Then
        cmdFont(1).Enabled = True
        Me.cmdWin.Enabled = True
    Else
        cmdFont(1).Enabled = False
        Me.cmdWin.Enabled = False
    End If
End Sub

Private Sub chk显示待发药_Click()
    If chk显示待发药.Value = 1 Then
        cmdFont(2).Enabled = True
        Me.cmdColor.Enabled = True
        Me.txt待发药行数.Enabled = True
        Me.txt待发药列数.Enabled = True
    Else
        cmdFont(2).Enabled = False
        Me.cmdColor.Enabled = False
        Me.txt待发药行数.Enabled = False
        Me.txt待发药列数.Enabled = False
    End If
End Sub

Private Sub chk显示待配药_Click()
    If chk显示待配药.Value = 1 Then
        cmdFont(3).Enabled = True
        Me.cmdPreparingColor.Enabled = True
        Me.txt待配药行数.Enabled = True
        Me.txt待配药列数.Enabled = True
    Else
        cmdFont(3).Enabled = False
        Me.cmdPreparingColor.Enabled = False
        Me.txt待配药行数.Enabled = False
        Me.txt待配药列数.Enabled = False
    End If
End Sub

Private Sub chk显示其他区域_Click()
    If chk显示其他区域.Value = 1 Then
        cmdFont(4).Enabled = True
        Me.cmdOther.Enabled = True
    Else
        cmdFont(4).Enabled = True
        Me.cmdOther.Enabled = True
    End If
End Sub

Private Sub cmdCall_Click()
    dlgColor.Color = shpCall.FillColor
    dlgColor.ShowColor
    shpCall.FillColor = dlgColor.Color
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click(Index As Integer)
    Dim strReg As String
    Dim str字体 As String
    
    On Error GoTo err
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    dlgFont(Index).Flags = cdlCFBoth
    dlgFont(Index).CancelError = False  '把点取消当作错误处理
'    dlgFont(Index).FontName = GetSetting("ZLSOFT", strReg, "字体(" & Index & ")", "宋体")
'    dlgFont(Index).FontBold = GetSetting("ZLSOFT", strReg, "粗体(" & Index & ")", "False")
'    dlgFont(Index).FontItalic = GetSetting("ZLSOFT", strReg, "斜体(" & Index & ")", "False")
'    dlgFont(Index).FontSize = GetSetting("ZLSOFT", strReg, "字号(" & Index & ")", "14")
    dlgFont(Index).ShowFont
    On Error GoTo 0
    '设置字体
    SaveSetting "ZLSOFT", strReg, "字体(" & Index & ")", dlgFont(Index).FontName
    SaveSetting "ZLSOFT", strReg, "粗体(" & Index & ")", dlgFont(Index).FontBold
    SaveSetting "ZLSOFT", strReg, "斜体(" & Index & ")", dlgFont(Index).FontItalic
    SaveSetting "ZLSOFT", strReg, "字号(" & Index & ")", dlgFont(Index).FontSize
    Me.lbl字体(Index) = dlgFont(Index).FontName & "," & IIf(dlgFont(Index).FontBold, "粗体,", "") & IIf(dlgFont(Index).FontItalic, "斜体,", "") & dlgFont(Index).FontSize
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strReg As String
    Dim strWin As String
    Dim i As Integer
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    SaveSetting "ZLSOFT", strReg, "左", Val(txtRect(1).Text)
    SaveSetting "ZLSOFT", strReg, "顶", Val(txtRect(2).Text)
    SaveSetting "ZLSOFT", strReg, "宽度", Val(txtRect(3).Text)
    SaveSetting "ZLSOFT", strReg, "高度", Val(txtRect(4).Text)
    
    SaveSetting "ZLSOFT", strReg, "窗口模式", IIf(Me.Opt窗口(0).Value = True, 0, 1)
    
    
    For i = 0 To Me.lst叫号窗口.ListCount - 1
        If lst叫号窗口.Selected(i) Then
            strWin = strWin & "," & lst叫号窗口.List(i)
        End If
    Next
    strWin = Mid(strWin, 2)
    SaveSetting "ZLSOFT", strReg, "窗口", strWin
    
    SaveSetting "ZLSOFT", strReg, "呼叫中颜色", shpCall.FillColor
    
    SaveSetting "ZLSOFT", strReg, "显示窗口", Me.chk显示窗体.Value
    SaveSetting "ZLSOFT", strReg, "窗口颜色", shpWin.FillColor
    
    SaveSetting "ZLSOFT", strReg, "显示其他内容", Me.chk显示其他区域.Value
    SaveSetting "ZLSOFT", strReg, "其他内容颜色", shpOther.FillColor
    
    SaveSetting "ZLSOFT", strReg, "显示待发药", Me.chk显示待发药.Value
    SaveSetting "ZLSOFT", strReg, "待发药人数", Me.lbl待发药Sum.Caption
    SaveSetting "ZLSOFT", strReg, "待发药行数", Me.txt待发药行数.Text
    SaveSetting "ZLSOFT", strReg, "待发药列数", Me.txt待发药列数.Text
    SaveSetting "ZLSOFT", strReg, "待发药颜色", shpCalling.FillColor
    
    SaveSetting "ZLSOFT", strReg, "显示待配药", Me.chk显示待配药.Value
    SaveSetting "ZLSOFT", strReg, "待配药人数", Me.lbl待发药Sum.Caption
    SaveSetting "ZLSOFT", strReg, "待配药行数", Me.txt待配药行数.Text
    SaveSetting "ZLSOFT", strReg, "待配药列数", Me.txt待配药列数.Text
    SaveSetting "ZLSOFT", strReg, "待配药颜色", shpPreparing.FillColor
    
    SaveSetting "ZLSOFT", strReg, "翻页时间", Me.txt翻页时间.Text
    SaveSetting "ZLSOFT", strReg, "刷新时间", Me.txt轮询时间.Text
    
    SaveSetting "ZLSOFT", strReg, "显示内容", Me.txtContent.Text
    
    Unload Me
End Sub


Private Sub cmdOther_Click()
    dlgColor.Color = shpOther.FillColor
    dlgColor.ShowColor
    shpOther.FillColor = dlgColor.Color
End Sub

Private Sub cmdWin_Click()
    dlgColor.Color = shpWin.FillColor
    dlgColor.ShowColor
    shpWin.FillColor = dlgColor.Color
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Integer
    Dim strWin As String
    Dim Index  As Integer
    
    strReg = "公共模块\药房排队叫号\液晶电视"
    
    Me.Opt窗口(Val(GetSetting("ZLSOFT", strReg, "窗口模式", "0"))).Value = True
    
    strWin = GetSetting("ZLSOFT", strReg, "窗口", "")
    '加载窗口
    LoadWin
    '恢复选中窗口
    For i = 0 To Me.lst叫号窗口.ListCount - 1
        If InStr(1, strWin, lst叫号窗口.List(i)) > 0 Then
            lst叫号窗口.Selected(i) = True
        End If
    Next
    
    '加载屏幕信息
    txtRect(1).Text = GetSetting("ZLSOFT", strReg, "左", "1024")
    txtRect(2).Text = GetSetting("ZLSOFT", strReg, "顶", "0")
    txtRect(3).Text = GetSetting("ZLSOFT", strReg, "宽度", "1024")
    txtRect(4).Text = GetSetting("ZLSOFT", strReg, "高度", "768")
    
    
    shpCall.FillColor = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
    
    Me.chk显示窗体.Value = GetSetting("ZLSOFT", strReg, "显示窗口", 1)
    shpWin.FillColor = GetSetting("ZLSOFT", strReg, "窗口颜色", vbGreen)
    
    Me.chk显示其他区域.Value = GetSetting("ZLSOFT", strReg, "显示其他内容", 1)
    shpOther.FillColor = GetSetting("ZLSOFT", strReg, "其他内容颜色", vbGreen)
    
    Me.chk显示待发药.Value = Val(GetSetting("ZLSOFT", strReg, "显示待发药", "1"))
    Me.txt待发药行数.Text = Val(GetSetting("ZLSOFT", strReg, "待发药行数", "3"))
    Me.txt待发药列数.Text = Val(GetSetting("ZLSOFT", strReg, "待发药列数", "3"))
    Me.lbl待发药Sum.Caption = Val(GetSetting("ZLSOFT", strReg, "待发药人数", "9"))
    shpCalling.FillColor = GetSetting("ZLSOFT", strReg, "待发药颜色", vbGreen)
    
    Me.chk显示待配药.Value = Val(GetSetting("ZLSOFT", strReg, "显示待配药", "1"))
    Me.txt待配药行数.Text = Val(GetSetting("ZLSOFT", strReg, "待配药行数", "3"))
    Me.txt待配药列数.Text = Val(GetSetting("ZLSOFT", strReg, "待配药列数", "3"))
    Me.lbl待配药Sum.Caption = Val(GetSetting("ZLSOFT", strReg, "待配药人数", "9"))
    shpPreparing.FillColor = GetSetting("ZLSOFT", strReg, "待配药颜色", vbGreen)
    
    Me.txt翻页时间.Text = GetSetting("ZLSOFT", strReg, "翻页时间", "5")
    Me.txt轮询时间.Text = GetSetting("ZLSOFT", strReg, "刷新时间", "10")
    Me.txtContent.Text = GetSetting("ZLSOFT", strReg, "显示内容", "")
    
    For Index = 0 To Me.dlgFont.UBound
        dlgFont(Index).Flags = cdlCFBoth
        dlgFont(Index).CancelError = False  '把点取消当作错误处理
        dlgFont(Index).FontName = GetSetting("ZLSOFT", strReg, "字体(" & Index & ")", "宋体")
        dlgFont(Index).FontBold = GetSetting("ZLSOFT", strReg, "粗体(" & Index & ")", "False")
        dlgFont(Index).FontItalic = GetSetting("ZLSOFT", strReg, "斜体(" & Index & ")", "False")
        dlgFont(Index).FontSize = GetSetting("ZLSOFT", strReg, "字号(" & Index & ")", "14")
        Me.lbl字体(Index) = dlgFont(Index).FontName & "," & IIf(dlgFont(Index).FontBold, "粗体,", "") & IIf(dlgFont(Index).FontItalic, "斜体,", "") & dlgFont(Index).FontSize
    Next

End Sub

Public Function ShowMe(ByVal strWins As String, ByVal frmParent As Form) As Boolean
'参数说明：strWins窗口串，格式为“窗口1,窗口2”
    mstrWins = strWins
    
    Me.Show 1, frmParent
    
    ShowMe = True
End Function

Private Sub lbl待发药行人数_Click()

End Sub

Private Sub Opt窗口_Click(Index As Integer)
    Me.fra叫号窗口.Enabled = IIf(Index = 0, False, True)
End Sub

Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待发药人数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待配药人数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待发药行数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待发药行数_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.txt待发药行数.Text <> "" Then
        Me.lbl待发药Sum.Caption = Val(Me.txt待发药行数.Text) * Val(Me.txt待发药列数.Text)
    End If
End Sub

Private Sub txt待发药列数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待发药列数_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txt待配药列数.Text = Me.txt待发药列数.Text
    
    If Me.txt待发药行数.Text <> "" Then
        Me.lbl待发药Sum.Caption = Val(Me.txt待发药行数.Text) * Val(Me.txt待发药列数.Text)
    End If
    
    If Me.txt待配药行数.Text <> "" Then
        Me.lbl待配药Sum.Caption = Val(Me.txt待配药行数.Text) * Val(Me.txt待配药列数.Text)
    End If
End Sub

Private Sub txt待配药行数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待配药行数_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.txt待配药行数.Text <> "" Then
        Me.lbl待配药Sum.Caption = Val(Me.txt待配药行数.Text) * (Me.txt待配药列数.Text)
    End If
End Sub

Private Sub txt待配药列数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt待配药列数_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txt待发药列数.Text = Me.txt待配药列数.Text
    
    If Me.txt待发药行数.Text <> "" Then
        Me.lbl待发药Sum.Caption = Val(Me.txt待发药行数.Text) * Val(Me.txt待发药列数.Text)
    End If
    
    If Me.txt待配药行数.Text <> "" Then
        Me.lbl待配药Sum.Caption = Val(Me.txt待配药行数.Text) * Val(Me.txt待配药列数.Text)
    End If
End Sub

Private Sub txt轮询时间_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt人数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cmdPreparingColor_Click()
    dlgColor.Color = shpPreparing.FillColor
    dlgColor.ShowColor
    shpPreparing.FillColor = dlgColor.Color
End Sub


Private Sub cmdColor_Click()
    dlgColor.Color = shpCalling.FillColor
    dlgColor.ShowColor
    shpCalling.FillColor = dlgColor.Color
End Sub

Private Sub LoadWin()
    Dim i As Integer
    
    For i = 0 To UBound(Split(mstrWins, ","))
        Me.lst叫号窗口.AddItem Split(mstrWins, ",")(i)
    Next
    
End Sub
