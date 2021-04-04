VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印设置"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   Icon            =   "frmPrintSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cbo打印机 
      Height          =   300
      Left            =   1275
      TabIndex        =   31
      Text            =   "cbo打印机"
      Top             =   2865
      Width           =   3660
   End
   Begin VB.TextBox txtPageSize 
      Height          =   300
      Index           =   1
      Left            =   3525
      TabIndex        =   29
      Top             =   3705
      Width           =   900
   End
   Begin VB.TextBox txtPageSize 
      Height          =   300
      Index           =   0
      Left            =   900
      TabIndex        =   27
      Top             =   3705
      Width           =   900
   End
   Begin VB.Frame Frame4 
      Height          =   135
      Left            =   135
      TabIndex        =   26
      Top             =   2655
      Width           =   4950
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   25
      Top             =   1275
      Width           =   4950
   End
   Begin VB.ComboBox cboPageSize 
      Height          =   300
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3210
      Width           =   3660
   End
   Begin VB.Frame Frame2 
      Caption         =   "方向"
      Height          =   1065
      Left            =   345
      TabIndex        =   20
      Top             =   1515
      Width           =   1320
      Begin VB.OptionButton opt方向 
         Caption         =   "横向"
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   22
         Top             =   690
         Width           =   885
      End
      Begin VB.OptionButton opt方向 
         Caption         =   "纵向"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   21
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "页边距(毫米)"
      Height          =   1065
      Left            =   1875
      TabIndex        =   11
      Top             =   1515
      Width           =   2880
      Begin VB.TextBox txt边距 
         Height          =   300
         Index           =   3
         Left            =   2040
         TabIndex        =   18
         Top             =   615
         Width           =   600
      End
      Begin VB.TextBox txt边距 
         Height          =   300
         Index           =   2
         Left            =   735
         TabIndex        =   16
         Top             =   615
         Width           =   600
      End
      Begin VB.TextBox txt边距 
         Height          =   300
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Top             =   270
         Width           =   600
      End
      Begin VB.TextBox txt边距 
         Height          =   300
         Index           =   0
         Left            =   735
         TabIndex        =   12
         Top             =   270
         Width           =   600
      End
      Begin VB.Label lbl边距 
         Caption         =   "右(R)"
         Height          =   195
         Index           =   3
         Left            =   1515
         TabIndex        =   19
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl边距 
         Caption         =   "左(L)"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl边距 
         Caption         =   "下(B)"
         Height          =   195
         Index           =   1
         Left            =   1515
         TabIndex        =   15
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lbl边距 
         Caption         =   "上(T)"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.OptionButton opt页码 
      Caption         =   "右"
      Height          =   225
      Index           =   2
      Left            =   3570
      TabIndex        =   10
      Top             =   1020
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton opt页码 
      Caption         =   "中"
      Height          =   225
      Index           =   1
      Left            =   2925
      TabIndex        =   9
      Top             =   1020
      Width           =   615
   End
   Begin VB.OptionButton opt页码 
      Caption         =   "左"
      Height          =   225
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      Top             =   1020
      Width           =   615
   End
   Begin VB.CheckBox chk页码 
      Caption         =   "打印页码"
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   1020
      Width           =   1125
   End
   Begin VB.TextBox txt标题Font 
      Enabled         =   0   'False
      Height          =   300
      Left            =   945
      TabIndex        =   5
      Top             =   555
      Width           =   3630
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "…"
      Height          =   300
      Left            =   4605
      TabIndex        =   4
      Top             =   555
      Width           =   300
   End
   Begin MSComDlg.CommonDialog cmdigFont 
      Left            =   5010
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txt标题 
      Height          =   300
      Left            =   945
      TabIndex        =   2
      Top             =   210
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3795
      TabIndex        =   1
      Top             =   4155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   0
      Top             =   4155
      Width           =   1100
   End
   Begin VB.Label Label3 
      Caption         =   "打印机(D)"
      Height          =   195
      Left            =   195
      TabIndex        =   32
      Top             =   2895
      Width           =   1050
   End
   Begin VB.Label lblPageSize 
      Caption         =   "高度(H)            毫米"
      Height          =   255
      Index           =   1
      Left            =   2820
      TabIndex        =   30
      Top             =   3750
      Width           =   2070
   End
   Begin VB.Label lblPageSize 
      Caption         =   "宽度(W)            毫米"
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   28
      Top             =   3750
      Width           =   2070
   End
   Begin VB.Label Label2 
      Caption         =   "纸张大小(Z)"
      Height          =   195
      Left            =   195
      TabIndex        =   24
      Top             =   3255
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "标题字体"
      Height          =   270
      Left            =   165
      TabIndex        =   6
      Top             =   615
      Width           =   870
   End
   Begin VB.Label lbl标题 
      Caption         =   "打印标题"
      Height          =   270
      Left            =   165
      TabIndex        =   3
      Top             =   255
      Width           =   750
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngIndex As Long
Const DC_PAPERS = 2          '表示要读取打印纸张大小
Const DC_PAPERNAMES = 16     '表示要读取打印纸张名称

'打印机的设备结构
Private Type DEVMODE
    dmdevicename As String * 64
    dmspecversion As Integer
    dmdriverversion As Integer
    dmsize As Integer
    dmdriverextra As Integer
    dmfields As Long
End Type

'API函数的声明?
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" ( _
ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, _
ByVal lpOutput As Long, lpDevMode As DEVMODE) As Long

Private Sub cboPageSize_Click()
    Dim intPage As Integer
    
    txtPageSize(0).Enabled = False
    txtPageSize(1).Enabled = False
    If cboPageSize.ListIndex >= 0 Then
        intPage = cboPageSize.ItemData(cboPageSize.ListIndex)
        If intPage > 0 And intPage < 256 Then
            vsPrint.vp.PaperSize = intPage
            txtPageSize(0).Text = Format(vsPrint.vp.PageWidth / (1440 / 25.4), "0.00")
            txtPageSize(1).Text = Format(vsPrint.vp.PageHeight / (1440 / 25.4), "0.00")
        ElseIf intPage = 256 Then
            vsPrint.vp.PaperSize = intPage
            txtPageSize(0).Text = Format(vsPrint.vp.PageWidth / (1440 / 25.4), "0.00")
            txtPageSize(1).Text = Format(vsPrint.vp.PageHeight / (1440 / 25.4), "0.00")
            txtPageSize(0).Enabled = True
            txtPageSize(1).Enabled = True
        End If
    End If
End Sub


Private Sub cbo打印机_Click()
    If cbo打印机.ListIndex >= 0 Then
        vsPrint.vp.Device = cbo打印机.List(cbo打印机.ListIndex)
        Call ReadSetup
    End If
End Sub

Private Sub chk页码_Click()
    If chk页码.Value = 1 Then
        opt页码(0).Enabled = True
        opt页码(1).Enabled = True
        opt页码(2).Enabled = True
    Else
        opt页码(0).Enabled = False
        opt页码(1).Enabled = False
        opt页码(2).Enabled = False
    End If
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click()
    Dim strFont As String
    
    strFont = Trim(txt标题Font)
    If strFont = "" Then strFont = "宋体|18"
    With cmdigFont
        .Flags = &H3 Or &H2 Or &H1 Or &H100
        .FontName = Split(strFont, "|")(0)
        .FontSize = Split(strFont, "|")(1)

        .ShowFont
        
        txt标题Font = .FontName & "|" & .FontSize
    End With
    
End Sub

Private Sub cmdOK_Click()
    If txtPageSize(0).Text <> "" Then
        If Not IsNumeric(txtPageSize(0).Text) Then
            MsgBox "页宽度错误！", vbInformation, Me.Caption
            txtPageSize(0).SetFocus
            Exit Sub
        End If
    End If
    
    If txtPageSize(1).Text <> "" Then
        If Not IsNumeric(txtPageSize(1).Text) Then
            MsgBox "页高度错误！", vbInformation, Me.Caption
            txtPageSize(1).SetFocus
            Exit Sub
        End If
    End If
    
    If txt边距(0).Text <> "" Then
        If Not IsNumeric(txt边距(0).Text) Then
            MsgBox "上边距错误！", vbInformation, Me.Caption
            txt边距(0).SetFocus
            Exit Sub
        End If
    End If
    
    If txt边距(1).Text <> "" Then
        If Not IsNumeric(txt边距(1).Text) Then
            MsgBox "下边距错误！", vbInformation, Me.Caption
            txt边距(1).SetFocus
            Exit Sub
        End If
    End If
    
    If txt边距(2).Text <> "" Then
        If Not IsNumeric(txt边距(2).Text) Then
            MsgBox "左边距错误！", vbInformation, Me.Caption
            txt边距(2).SetFocus
            Exit Sub
        End If
    End If
    
    If txt边距(3).Text <> "" Then
        If Not IsNumeric(txt边距(3).Text) Then
            MsgBox "右边距错误！", vbInformation, Me.Caption
            txt边距(3).SetFocus
            Exit Sub
        End If
    End If
    
    Call SaveSetup
    Unload Me
End Sub

Private Sub Form_Load()

    Call LoadPrint
End Sub

Public Sub PageSetup(ByVal Index As Long)
    mlngIndex = Index
    Me.Show vbModal
    
End Sub

Private Sub SaveSetup()
    Dim strValue As String
    
    If cbo打印机.ListIndex < 0 Then Exit Sub
    
    strValue = cbo打印机.List(cbo打印机.ListIndex)
    WriteIni "Report" & mlngIndex, "打印机", strValue, App.Path & "\PrintSetup.ini"
    
    strValue = txt标题
    WriteIni "Report" & mlngIndex, "标题", strValue, App.Path & "\PrintSetup.ini"
    strValue = txt标题Font
    WriteIni "Report" & mlngIndex, "标题字体", strValue, App.Path & "\PrintSetup.ini"
    
    If chk页码.Value = 1 Then
        If opt页码(0).Value = True Then
            WriteIni "Report" & mlngIndex, "打印页码", "1", App.Path & "\PrintSetup.ini"
        ElseIf opt页码(1).Value = True Then
            WriteIni "Report" & mlngIndex, "打印页码", "2", App.Path & "\PrintSetup.ini"
        ElseIf opt页码(2).Value = True Then
            WriteIni "Report" & mlngIndex, "打印页码", "3", App.Path & "\PrintSetup.ini"
        End If
    Else
        WriteIni "Report" & mlngIndex, "打印页码", "0", App.Path & "\PrintSetup.ini"
    End If
    
    'vp.Orientation
    If opt方向(0).Value = True Then
        WriteIni "Report" & mlngIndex, "打印方向", "0", App.Path & "\PrintSetup.ini"
    ElseIf opt方向(1).Value = True Then
        WriteIni "Report" & mlngIndex, "打印方向", "1", App.Path & "\PrintSetup.ini"
    End If
    
    strValue = Val(txt边距(0).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "上边距", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "上边距", "25.4", App.Path & "\PrintSetup.ini"
    End If
    
    strValue = Val(txt边距(1).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "下边距", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "下边距", "25.4", App.Path & "\PrintSetup.ini"
    End If
    
    strValue = Val(txt边距(2).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "左边距", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "左边距", "25.4", App.Path & "\PrintSetup.ini"
    End If

    strValue = Val(txt边距(3).Text)
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        WriteIni "Report" & mlngIndex, "右边距", Val(strValue), App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "右边距", "25.4", App.Path & "\PrintSetup.ini"
    End If

    strValue = cboPageSize.ItemData(cboPageSize.ListIndex)
    
    If Val(strValue) <> 256 Then
        vsPrint.vp.PaperSize = Val(strValue)
        WriteIni "Report" & mlngIndex, "纸张大小", Val(strValue), App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "纸张宽度", vsPrint.vp.PageWidth, App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "纸张高度", vsPrint.vp.PageHeight, App.Path & "\PrintSetup.ini"
    Else
        WriteIni "Report" & mlngIndex, "纸张大小", Val(strValue), App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "纸张宽度", Val(txtPageSize(0).Text) * (1440 / 25.4), App.Path & "\PrintSetup.ini"
        WriteIni "Report" & mlngIndex, "纸张高度", Val(txtPageSize(1).Text) * (1440 / 25.4), App.Path & "\PrintSetup.ini"
    End If
End Sub

Private Sub LoadPrint()
    Dim i As Integer, strValue As String
    
    cbo打印机.Clear
    For i = 0 To vsPrint.vp.NDevices - 1
        cbo打印机.AddItem vsPrint.vp.Devices(i)
    Next
    
    If cbo打印机.ListCount <= 0 Then
        MsgBox "请安装打印机后再使用此功能！", vbInformation, Me.Caption
        Unload Me
        Exit Sub
    End If
    
    strValue = ReadIni("Report" & mlngIndex, "打印机", App.Path & "\PrintSetup.ini")

    For i = 0 To cbo打印机.ListCount - 1
        If strValue = cbo打印机.List(i) Then
            cbo打印机.ListIndex = i
            vsPrint.vp.Device = strValue
            Exit For
        End If
    Next
    
    If cbo打印机.ListIndex < 0 Then
        cbo打印机.ListIndex = 0
        vsPrint.vp.Device = cbo打印机.List(cbo打印机.ListIndex)
    End If
End Sub

Private Sub ReadSetup()

    Dim strValue As String, i As Integer
    
'    cboPageSize.Clear
'    For i = 1 To 256
'        If vsPrint.vp.PaperSizes(i) Then
'            strValue = getPageName(i)
'            If strValue <> "" Then
'                cboPageSize.AddItem strValue
'                cboPageSize.ItemData(cboPageSize.NewIndex) = i
'            End If
'        End If
'    Next
    Call FillPaperSize
    
    txt标题 = ReadIni("Report" & mlngIndex, "标题", App.Path & "\PrintSetup.ini")
    txt标题Font = ReadIni("Report" & mlngIndex, "标题字体", App.Path & "\PrintSetup.ini")
    
    strValue = ReadIni("Report" & mlngIndex, "打印页码", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 4 Then
        chk页码.Value = 1
        opt页码(Val(strValue) - 1).Value = True
        opt页码(0).Enabled = True
        opt页码(1).Enabled = True
        opt页码(2).Enabled = True
    Else
        chk页码.Value = 0
        opt页码(0).Enabled = False
        opt页码(1).Enabled = False
        opt页码(2).Enabled = False
    End If
    
    '方向
    strValue = ReadIni("Report" & mlngIndex, "打印方向", App.Path & "\PrintSetup.ini")
    'vp.Orientation
    
    If Val(strValue) = 0 Then
        opt方向(0).Value = True
    Else
        opt方向(1).Value = True
    End If
    
    '边距
    strValue = ReadIni("Report" & mlngIndex, "上边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt边距(0).Text = Format(Val(strValue), "0.0")
    Else
        txt边距(0).Text = "25.4"
    End If
    strValue = ReadIni("Report" & mlngIndex, "下边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt边距(1).Text = Format(Val(strValue), "0.0")
    Else
        txt边距(1).Text = "25.4"
    End If
    strValue = ReadIni("Report" & mlngIndex, "左边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt边距(2).Text = Format(Val(strValue), "0.0")
    Else
        txt边距(2).Text = "25.4"
    End If
    strValue = ReadIni("Report" & mlngIndex, "右边距", App.Path & "\PrintSetup.ini")
    If Val(strValue) > 0 And Val(strValue) < 100 Then
        txt边距(3).Text = Format(Val(strValue), "0.0")
    Else
        txt边距(3).Text = "25.4"
    End If
    
    '纸张大小
    strValue = ReadIni("Report" & mlngIndex, "纸张大小", App.Path & "\PrintSetup.ini")
    For i = 0 To cboPageSize.ListCount - 1
        If cboPageSize.ItemData(i) = Val(strValue) Then
            cboPageSize.ListIndex = i
            Exit For
        End If
    Next
    
    If cboPageSize.ListIndex < 0 Then
        For i = 0 To cboPageSize.ListCount - 1
            If cboPageSize.ItemData(i) = 9 Then
                cboPageSize.ListIndex = i 'A4
                Exit For
            End If
        Next
    End If
    '纸张宽高
    If Val(strValue) <> 256 Then
        If Val(strValue) > 0 And Val(strValue) <= 256 Then vsPrint.vp.PaperSize = Val(strValue)
    Else
        strValue = ReadIni("Report" & mlngIndex, "纸张宽度", App.Path & "\PrintSetup.ini")
        txtPageSize(0).Text = Val(strValue) / (1440 / 25.4)
        strValue = ReadIni("Report" & mlngIndex, "纸张高度", App.Path & "\PrintSetup.ini")
        txtPageSize(1).Text = Val(strValue) / (1440 / 25.4)
    End If
End Sub

Private Sub FillPaperSize()

    '根据当前使用的打印机取得所有打印纸张和大小填充到ComBoBox显示控件中(cboPaper)。
    
    On Error Resume Next
    
    Dim devname As String
    Dim devoutput As String
    Dim papercount As Long
    Dim bytepapernames() As Byte
    Dim bytepapersizes() As Byte
    Dim sinfo As String
    Dim X As Long
    Dim di As Long
    Dim spapersizes As String
    Dim dv As DEVMODE
    
    Screen.MousePointer = vbHourglass
    
    devname = cbo打印机.List(cbo打印机.ListIndex)   ' DeviceName
    
    '当前打印机的名称
    
    devoutput = vsPrint.vp.Port
    
    '当前打印机的输出端口名称得到当前打印机的打印纸张数?
    
    papercount = DeviceCapabilities(devname, devoutput, DC_PAPERNAMES, 0&, dv)
    
    If papercount = 0 Then
    
        'MsgBox "该打印机无有效的打印纸张大小?", vbInformation, "打印设置"
        Exit Sub
    
    End If
    
    '为保存打印用纸名称申请空间
    
    ReDim bytepapernames(1 To 64 * papercount)
    
    '一个纸张名称需要64个字符空间来存储
    
    '取出打印机上的所有用纸名称
    
    DeviceCapabilities devname, devoutput, DC_PAPERNAMES, VarPtr(bytepapernames(1)), dv
    
    '为保存打印用纸所对应的PaperSize的字符串申请空间
    
    ReDim bytepapersizes(1 To 2 * papercount)
    
    '一个PaperSize需要2 个字符空间来存储取出打印机上的所有用纸所对应的PaperSize
    
    DeviceCapabilities devname, devoutput, DC_PAPERS, VarPtr(bytepapersizes(1)), dv
    
    cboPageSize.Clear
    
    '为了正确的取汉字?用StrConv方法对PaperName进行转换?
    
    For X = 1 To papercount
    
    '一次取出一个打印用纸的名称
    
        sinfo = StrConv(MidB(bytepapernames, (X - 1) * 64 + 1, 64), vbUnicode)
        
        cboPageSize.AddItem Left(sinfo, InStr(sinfo, Chr(0)) - 1)  '纸张名称
        
        cboPageSize.ItemData(cboPageSize.NewIndex) = bytepapersizes((X - 1) * 2 + 1)  '纸张大小
    
    Next X
    
    'If cboPageSize.ListCount > 0 Then cboPageSize.ListIndex = 0
    
    Screen.MousePointer = vbDefault

End Sub
