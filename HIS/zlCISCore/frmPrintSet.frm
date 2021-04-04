VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印设置"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame5 
      Caption         =   "边距(mm)"
      Height          =   1065
      Left            =   120
      TabIndex        =   30
      Top             =   2265
      Width           =   2385
      Begin VB.TextBox txt右 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1455
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   600
         Width           =   540
      End
      Begin MSComCtl2.UpDown UD下 
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt下"
         BuddyDispid     =   196611
         OrigLeft        =   3750
         OrigTop         =   255
         OrigRight       =   3990
         OrigBottom      =   525
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD上 
         Height          =   315
         Left            =   915
         TabIndex        =   7
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt上"
         BuddyDispid     =   196612
         OrigLeft        =   2385
         OrigTop         =   240
         OrigRight       =   2625
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD左 
         Height          =   315
         Left            =   915
         TabIndex        =   11
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt左"
         BuddyDispid     =   196613
         OrigLeft        =   1080
         OrigTop         =   240
         OrigRight       =   1320
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt下 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1455
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   270
         Width           =   540
      End
      Begin VB.TextBox txt上 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   270
         Width           =   540
      End
      Begin VB.TextBox txt左 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   615
         Width           =   525
      End
      Begin MSComCtl2.UpDown UD右 
         Height          =   300
         Left            =   2010
         TabIndex        =   13
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt右"
         BuddyDispid     =   196610
         OrigLeft        =   1080
         OrigTop         =   240
         OrigRight       =   1320
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右"
         Height          =   180
         Left            =   1245
         TabIndex        =   36
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下"
         Height          =   180
         Left            =   1245
         TabIndex        =   33
         Top             =   330
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上"
         Height          =   180
         Left            =   150
         TabIndex        =   32
         Top             =   330
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左"
         Height          =   180
         Left            =   150
         TabIndex        =   31
         Top             =   675
         Width           =   180
      End
   End
   Begin VB.Frame fraOrient 
      Caption         =   "纸向"
      Height          =   1065
      Left            =   2520
      TabIndex        =   34
      Top             =   2265
      Width           =   1425
      Begin VB.OptionButton opt纵向 
         Caption         =   "纵向"
         Height          =   285
         Left            =   675
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton opt横向 
         Caption         =   "横向"
         Height          =   285
         Left            =   675
         TabIndex        =   15
         Top             =   600
         Width           =   660
      End
      Begin VB.Image img纵向 
         Height          =   480
         Left            =   120
         Picture         =   "frmPrintSet.frx":058A
         Top             =   330
         Width           =   480
      End
      Begin VB.Image img横向 
         Height          =   480
         Left            =   120
         Picture         =   "frmPrintSet.frx":0E54
         Top             =   330
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "打印机"
      Height          =   1005
      Left            =   120
      TabIndex        =   27
      Top             =   90
      Width           =   5850
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   1185
         TabIndex        =   29
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "位置"
         Height          =   180
         Left            =   1185
         TabIndex        =   28
         Top             =   660
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmPrintSet.frx":171E
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "纸张"
      Height          =   1065
      Left            =   120
      TabIndex        =   21
      Top             =   1155
      Width           =   3825
      Begin VB.ComboBox cboPage 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2955
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   3
         TabIndex        =   2
         Top             =   630
         Width           =   480
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2415
         MaxLength       =   3
         TabIndex        =   4
         Top             =   630
         Width           =   480
      End
      Begin MSComCtl2.UpDown UDHeight 
         Height          =   285
         Left            =   2895
         TabIndex        =   5
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196631
         OrigLeft        =   2985
         OrigTop         =   630
         OrigRight       =   3225
         OrigBottom      =   930
         Max             =   460
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDWidth 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196630
         OrigLeft        =   1200
         OrigTop         =   645
         OrigRight       =   1440
         OrigBottom      =   945
         Max             =   460
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大小"
         Height          =   180
         Left            =   285
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宽度"
         Height          =   180
         Left            =   300
         TabIndex        =   25
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高度"
         Height          =   180
         Left            =   2010
         TabIndex        =   24
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1515
         TabIndex        =   23
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   3210
         TabIndex        =   22
         Top             =   690
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4455
      TabIndex        =   17
      Top             =   1215
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4455
      TabIndex        =   18
      Top             =   1665
      Width           =   1100
   End
   Begin VB.Frame Frame4 
      Caption         =   "送纸器"
      Height          =   765
      Left            =   120
      TabIndex        =   19
      Top             =   3375
      Width           =   3825
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "纸张来源"
         Height          =   180
         Left            =   210
         TabIndex        =   20
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   3975
      ScaleHeight     =   491.128
      ScaleMode       =   0  'User
      ScaleWidth      =   491.128
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2055
      Width           =   2130
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   405
         ScaleHeight     =   1455
         ScaleMode       =   0  'User
         ScaleWidth      =   1140
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   450
         ScaleHeight     =   1485
         ScaleWidth      =   1170
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   315
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnWinNT As Boolean
Private mdblW As Double  '左边不可打印比例
Private mdblH As Double  '上边不可打印比例

'打印参数变量
Private mstrPrinter As String '打印机
Private mintPage As Integer '纸张
Private mlngWidth As Long '自定义纸张宽度,Twip
Private mlngHeight As Long '自定义纸张高度'Twip
Private mintOrient As Integer   '纸向
Private mintBin As Integer '进纸方式
Private mlngLeft As Long '左边距'mm
Private mlngRight As Long '右边距'mm
Private mlngTop As Long '上边距'mm
Private mlngBottom As Long '下边距'mm

'事件控制
Private mblnChange As Boolean

Private Sub cboBin_Click()
    If cboBin.ListIndex <> -1 Then
        mintBin = cboBin.ItemData(cboBin.ListIndex)
    End If
End Sub

Private Sub cboPage_Click()
    Dim blnOK As Boolean
    Dim dblRight As Double
    Dim dblDown As Double
    
    
    '纸张
    If cboPage.ItemData(cboPage.ListIndex) <> 256 Then
        Printer.PaperSize = cboPage.ItemData(cboPage.ListIndex)
        mintPage = Printer.PaperSize
    Else
        '强行设置自定义纸张可用,不检查
        mintPage = 256
    End If
        
    '纸向
    If mintPage <> 256 Then
        On Error Resume Next
        Printer.Orientation = mintOrient
        mintOrient = Printer.Orientation
    Else
        mintOrient = 1
    End If
    On Error GoTo 0
    fraOrient.Enabled = mintPage <> 256
        
    '最后实际设置纸张大小(纸向影响之后)
    If mintPage <> 256 Then
        '取该打印机支持该幅面的真实尺寸
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
        
        '不可打印区域比例
        mdblW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        mdblH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    Else
        '自定义纸张认为全部可以打印
        mdblW = 0
        mdblH = 0
    End If
    
    '显示纸张尺寸
    mblnChange = False
    txtWidth.Tag = mlngWidth
    txtWidth.Text = CLng(mlngWidth / 56.7)
    txtHeight.Tag = mlngHeight
    txtHeight.Text = CLng(mlngHeight / 56.7)
    mblnChange = True
    
    '显示可用边距
    '最小在可打印区域之内
    '最大不超过宽高的1/4
    UD左.Min = mlngWidth / 56.7 * mdblW
    UD左.Max = mlngWidth / 56.7 / 4
    UD右.Min = UD左.Min
    UD右.Max = UD左.Max
    
    UD上.Min = mlngHeight / 56.7 * mdblH
    UD上.Max = mlngHeight / 56.7 / 4
    UD下.Min = UD上.Min
    UD下.Max = UD上.Max
    
    If mlngLeft >= UD左.Min And mlngLeft <= UD左.Max Then
        UD左.Value = mlngLeft
    Else
        UD左.Value = UD左.Min
    End If
    If mlngRight >= UD右.Min And mlngRight <= UD右.Max Then
        UD右.Value = mlngRight
    Else
        UD右.Value = UD右.Min
    End If
    If mlngTop >= UD上.Min And mlngTop <= UD上.Max Then
        UD上.Value = mlngTop
    Else
        UD上.Value = UD上.Min
    End If
    If mlngBottom >= UD下.Min And mlngBottom <= UD下.Max Then
        UD下.Value = mlngBottom
    Else
        UD下.Value = UD下.Min
    End If
    
    mlngLeft = UD左.Value
    mlngRight = UD右.Value
    mlngTop = UD上.Value
    mlngBottom = UD下.Value
    
    '显示纸向
    mblnChange = False
    If mintOrient = 1 Then
        opt纵向.Value = True: opt纵向_Click
    Else
        opt横向.Value = True: opt横向_Click
    End If
    mblnChange = True
    
    '显示预览纸张
    Call ShowPaper
End Sub

Private Sub cboPrinter_Click()
    Dim i As Integer, j As Integer
    Dim lngCount As Long, strTmp As String
    Dim strPaperSize As String * 300
    Dim strPaperBin As String * 100
    Dim strPaperBinName As String * 1000
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    mstrPrinter = Printer.DeviceName
    lblLoc.Caption = "位置: " & Printer.Port
    
    '如果支持,则保持原有纸张
    If mintPage <> 256 Then
        On Error Resume Next
        Printer.PaperSize = mintPage
        On Error GoTo 0
        mintPage = Printer.PaperSize
    End If
    
    '设置可用纸张
    '返回格式见常量说明或MSDN
    cboPage.Clear
    '------------------------------------------------------------------------------------------
    '纸张大小
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaperSize, 0)
    For i = 1 To lngCount
        j = Asc(Mid(strPaperSize, i * 2, 1)) * 256 + Asc(Mid(strPaperSize, i * 2 - 1, 1))
        If j >= 1 And j <= 41 Then '只列出标准支持的纸张
            cboPage.AddItem GetPaperName(j)
            cboPage.ItemData(cboPage.ListCount - 1) = j
            If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
        End If
    Next
    '------------------------------------------------------------------------------------------
    '自定义纸张处理
    i = 256
    cboPage.AddItem GetPaperName(i)
    cboPage.ItemData(cboPage.ListCount - 1) = i
    If mintPage = 256 Then cboPage.ListIndex = cboPage.NewIndex
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '如果支持,则保持原有进纸方式
    On Error Resume Next
    Printer.PaperBin = mintBin
    On Error GoTo 0
    mintBin = Printer.PaperBin
    
    '设置可用进纸方式
    cboBin.Clear
    '------------------------------------------------------------------------------------------
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
    j = 1
    For i = 1 To lngCount
        '进纸名称
        Do
            If Mid(strPaperBinName, j, 1) = Chr(0) Then
                cboBin.AddItem Trim(strTmp)
                
                '进纸编号
                cboBin.ItemData(cboBin.ListCount - 1) = Asc(Mid(strPaperBin, i * 2, 1)) * 256 + Asc(Mid(strPaperBin, i * 2 - 1, 1))
                If cboBin.ItemData(cboBin.ListCount - 1) = mintBin Then
                    cboBin.ListIndex = cboBin.NewIndex
                End If

                j = 24 + j - LenB(StrConv(strTmp, vbFromUnicode))
                strTmp = ""
                Exit Do
            Else
                strTmp = strTmp & Mid(strPaperBinName, j, 1)
                j = j + 1
            End If
        Loop
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "请确定报表的纸张宽度！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    If CInt(txtWidth.Text) > UDWidth.Max Then
        MsgBox "报表的纸张宽度不能超过" & UDWidth.Max & "毫米！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "请确定报表的纸张高度！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Sub
    End If
    If CInt(txtHeight.Text) > UDHeight.Max Then
        MsgBox "报表的纸张高度不能超过" & UDHeight.Max & "毫米！", vbExclamation, App.Title
        txtHeight.SetFocus: Exit Sub
    End If
    
    '保存打印参数
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "打印机", mstrPrinter
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸张", mintPage
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", mlngWidth
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", mlngHeight
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸向", mintOrient
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "进纸", mintBin
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", mlngLeft
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", mlngRight
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", mlngTop
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", mlngBottom
    
    gblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    If Printers.Count = 0 Then
        MsgBox "系统中没有安装任何打印机,请先安装打印机！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    gblnOK = False
    mblnChange = True
    
    mblnWinNT = IsWindowsNT
    
    '初始化打印参数
    mstrPrinter = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "打印机", Printer.DeviceName)
    mintPage = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)
    mlngWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "宽度", Printer.Width)
    mlngHeight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "高度", Printer.Height)
    mintOrient = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸向", Printer.Orientation)
    mintBin = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "进纸", Printer.PaperBin)
    mlngLeft = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "左边距", OFFSET_LEFT)
    mlngRight = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "右边距", OFFSET_RIGHT)
    mlngTop = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "上边距", OFFSET_TOP)
    mlngBottom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "下边距", OFFSET_BOTTOM)
    
    '初始打印机列表
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '打印机索引
            
            '读取存储的打印机为当前打印机,并初始化可用页面
            If mstrPrinter = Printers(i).DeviceName Then .ListIndex = .NewIndex
        Next
        
        '缺省初始化为当前打印机
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '读取系统当前的打印机为当前打印机,并初始化可用页面
                If .List(i) = Printer.DeviceName Then .ListIndex = i: Exit For
            Next
        End If
    End With
    
    '边距
    txt左.Text = mlngLeft
    txt右.Text = mlngRight
    txt上.Text = mlngTop
    txt下.Text = mlngBottom
End Sub

Private Sub opt横向_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If opt横向.Value Then
        img纵向.Visible = False
        img横向.Visible = True
        
        If mintOrient = 1 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
            
            mlngLeft = lngB
            mlngRight = lngT
            mlngTop = lngL
            mlngBottom = lngR
        End If
        
        mintOrient = 2
        
        If mblnChange Then Call cboPage_Click
    End If
End Sub

Private Sub opt纵向_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If opt纵向.Value Then
        img纵向.Visible = True
        img横向.Visible = False
        
        If mintOrient = 2 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
              
            mlngLeft = lngT
            mlngRight = lngB
            mlngTop = lngR
            mlngBottom = lngL
        End If
        
        mintOrient = 1
        
        If mblnChange Then Call cboPage_Click
    End If
End Sub

Private Sub txtHeight_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtHeight.Text) Then
        txtHeight.Tag = CLng(txtHeight.Text * 56.7)
        mlngHeight = CLng(txtHeight.Text * 56.7)
        
        cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
End Sub

Private Sub txtWidth_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtWidth.Text) Then
        txtWidth.Tag = CLng(txtWidth.Text * 56.7)
        mlngWidth = CLng(txtWidth.Text * 56.7)
        
        cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = 0: txtHeight.SelLength = Len(txtHeight.Text)
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
End Sub

Private Sub txt上_GotFocus()
    zlControl.TxtSelAll txt上
End Sub

Private Sub txt上_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt下_GotFocus()
    zlControl.TxtSelAll txt下
End Sub

Private Sub txt下_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt右_GotFocus()
    zlControl.TxtSelAll txt右
End Sub

Private Sub txt左_GotFocus()
    zlControl.TxtSelAll txt左
End Sub

Private Sub txt左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub UD上_Change()
    mlngTop = UD上.Value
    Call ShowPaper
End Sub

Private Sub UD下_Change()
    mlngBottom = UD下.Value
    Call ShowPaper
End Sub

Private Sub UD右_Change()
    mlngRight = UD右.Value
    Call ShowPaper
End Sub

Private Sub UD左_Change()
    mlngLeft = UD左.Value
    Call ShowPaper
End Sub

Private Sub ShowPaper()
'功能：显示设置的纸张的预览
    On Error Resume Next
    
    picPaper.Cls
    
    picPaper.Width = mlngWidth / 56.7
    picPaper.Height = mlngHeight / 56.7
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth
    picPaper.ScaleHeight = mlngHeight
    
    picPaper.Line (0, mlngTop * 56.7)-(picPaper.ScaleWidth, mlngTop * 56.7), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - (mlngBottom + 2) * 56.7)-(picPaper.ScaleWidth, picPaper.ScaleHeight - (mlngBottom + 2) * 56.7), &H808080
    
    picPaper.Line (mlngLeft * 56.7, 0)-(mlngLeft * 56.7, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - (mlngRight + 2) * 56.7, 0)-(picPaper.ScaleWidth - (mlngRight + 2) * 56.7, picPaper.ScaleHeight), &H808080
    
    Me.Refresh
End Sub
