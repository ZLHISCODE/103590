VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChildPrintSet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5400
      Index           =   0
      Left            =   300
      ScaleHeight     =   5400
      ScaleWidth      =   5880
      TabIndex        =   0
      Top             =   255
      Width           =   5880
      Begin VB.TextBox txt上 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   2595
         Width           =   465
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   930
         TabIndex        =   18
         Top             =   165
         Width           =   4815
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   3
         Left            =   930
         TabIndex        =   17
         Top             =   1320
         Width           =   4815
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   495
         Width           =   4335
      End
      Begin VB.ComboBox cboPage 
         Height          =   300
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1530
         Width           =   4335
      End
      Begin VB.TextBox txtWidth 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1890
         Width           =   780
      End
      Begin VB.TextBox txtHeight 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   3345
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1890
         Width           =   885
      End
      Begin VB.TextBox txt右 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   2730
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   2970
         Width           =   465
      End
      Begin VB.TextBox txt下 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   2730
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   2595
         Width           =   465
      End
      Begin VB.TextBox txt左 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   2970
         Width           =   465
      End
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4635
         Width           =   3960
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   0
         Left            =   930
         TabIndex        =   8
         Top             =   2295
         Width           =   4815
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   1
         Left            =   930
         TabIndex        =   7
         Top             =   3420
         Width           =   2805
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   4
         Left            =   930
         TabIndex        =   6
         Top             =   4395
         Width           =   4815
      End
      Begin VB.OptionButton opt纵向 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "纵向"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1575
         TabIndex        =   5
         Top             =   3825
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton opt横向 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "横向"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2415
         TabIndex        =   4
         Top             =   3825
         Width           =   660
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1710
         Left            =   3675
         ScaleHeight     =   394.286
         ScaleMode       =   0  'User
         ScaleWidth      =   491.128
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2505
         Width           =   2130
         Begin VB.PictureBox picPaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   495
            ScaleHeight     =   1455
            ScaleMode       =   0  'User
            ScaleWidth      =   1140
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   105
            Width           =   1170
         End
         Begin VB.PictureBox picShadow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   540
            ScaleHeight     =   1485
            ScaleWidth      =   1170
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   150
            Width           =   1170
         End
      End
      Begin MSComCtl2.UpDown UDHeight 
         Height          =   285
         Left            =   4230
         TabIndex        =   19
         Top             =   1875
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         OrigLeft        =   2985
         OrigTop         =   630
         OrigRight       =   3225
         OrigBottom      =   930
         Max             =   460
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UDWidth 
         Height          =   285
         Left            =   2130
         TabIndex        =   20
         Top             =   1875
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txt上"
         BuddyDispid     =   196610
         OrigLeft        =   1200
         OrigTop         =   645
         OrigRight       =   1440
         OrigBottom      =   945
         Max             =   460
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD下 
         Height          =   270
         Left            =   3195
         TabIndex        =   21
         Top             =   2595
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "opt纵向"
         BuddyDispid     =   196620
         OrigLeft        =   3750
         OrigTop         =   255
         OrigRight       =   3990
         OrigBottom      =   525
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD上 
         Height          =   270
         Left            =   1815
         TabIndex        =   22
         Top             =   2595
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   25
         OrigLeft        =   2385
         OrigTop         =   240
         OrigRight       =   2625
         OrigBottom      =   540
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD左 
         Height          =   270
         Left            =   1815
         TabIndex        =   23
         Top             =   2970
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "cboBin"
         BuddyDispid     =   196619
         OrigLeft        =   1080
         OrigTop         =   240
         OrigRight       =   1320
         OrigBottom      =   540
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD右 
         Height          =   300
         Left            =   3195
         TabIndex        =   24
         Top             =   2940
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "picBack"
         BuddyDispid     =   196622
         OrigLeft        =   1080
         OrigTop         =   240
         OrigRight       =   1320
         OrigBottom      =   540
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "打 印 机"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   46
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "纸    张"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   16
         Left            =   195
         TabIndex        =   45
         Top             =   1305
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "frmChildPrintSet.frx":0000
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   930
         TabIndex        =   44
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "位置"
         Height          =   180
         Left            =   930
         TabIndex        =   43
         Top             =   915
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大小"
         Height          =   180
         Left            =   930
         TabIndex        =   42
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宽度"
         Height          =   180
         Left            =   930
         TabIndex        =   41
         Top             =   1935
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高度"
         Height          =   180
         Left            =   2940
         TabIndex        =   40
         Top             =   1935
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Index           =   4
         Left            =   2415
         TabIndex        =   39
         Top             =   1935
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Index           =   3
         Left            =   4515
         TabIndex        =   38
         Top             =   1935
         Width           =   180
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右边"
         Height          =   180
         Left            =   2340
         TabIndex        =   37
         Top             =   3000
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下边"
         Height          =   180
         Left            =   2340
         TabIndex        =   36
         Top             =   2625
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上边"
         Height          =   180
         Left            =   930
         TabIndex        =   35
         Top             =   2625
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左边"
         Height          =   180
         Left            =   930
         TabIndex        =   34
         Top             =   3000
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "纸张来源"
         Height          =   180
         Left            =   930
         TabIndex        =   33
         Top             =   4695
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "边    距"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   32
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "纸张来源"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   31
         Top             =   4380
         Width           =   720
      End
      Begin VB.Image img纵向 
         Height          =   480
         Left            =   975
         Picture         =   "frmChildPrintSet.frx":08CA
         Top             =   3720
         Width           =   480
      End
      Begin VB.Image img横向 
         Height          =   480
         Left            =   975
         Picture         =   "frmChildPrintSet.frx":1194
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "纸    向"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   30
         Top             =   3405
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Index           =   5
         Left            =   3480
         TabIndex        =   29
         Top             =   2625
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Index           =   6
         Left            =   3465
         TabIndex        =   28
         Top             =   3000
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Index           =   7
         Left            =   2085
         TabIndex        =   27
         Top             =   2625
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Index           =   8
         Left            =   2085
         TabIndex        =   26
         Top             =   3000
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmChildPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const HORZRES = 8            '  Horizontal Width in pixels
Private Const VERTRES = 10           '  Vertical Width in pixels
Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90       '  Logical pixels/inch in Y
Private Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Private Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Private Const PHYSICALHEIGHT = 111  '  Physical Height in device units
Private Const PHYSICALWidth = 110   '  Physical Width in device units
Private Const DC_PAPERNAMES = 16    '纸张名称(每64字符为一段,以Chr(0)结束)
Private Const DC_PAPERS = 2         '纸张编号(Array or Word)
Private Const DC_BINNAMES = 12      '进纸方式(每24字符为一段,以Chr(0)结束)
Private Const DC_BINS = 6           '进纸编号(Array or Word)

Private Const OFFSET_LEFT = 20
Private Const OFFSET_TOP = 20
Private Const OFFSET_RIGHT = 20
Private Const OFFSET_BOTTOM = 20

Private mdblW As Double             '左边不可打印比例
Private mdblH As Double             '上边不可打印比例

'打印参数变量
Private mstrPrinter As String       '打印机
Private mintPage As Integer         '纸张
Private mstrPagerName As String     '纸张名称
Private mlngWidth As Long           '自定义纸张宽度,Twip
Private mlngHeight As Long          '自定义纸张高度'Twip
Private mintOrient As Integer       '纸向
Private mintBin As Integer          '进纸方式
Private mlngLeft As Long            '左边距'mm
Private mlngRight As Long           '右边距'mm
Private mlngTop As Long             '上边距'mm
Private mlngBottom As Long          '下边距'mm

'事件控制
Private mclsPrint As New clsPrint

Private mblnDataChanged As Boolean
Public mbytMode As Byte
Private mblnOK As Boolean
Private mblnModifyPaper As Boolean
Private mintPaperSize As Integer
Private mfrmMain As Object
Private mstrSavePath As String
Private mblnReading As Boolean

Public Event Activate()                                 '子窗体激活

'######################################################################################################################

Public Function GetPaper(ByRef objPaper As USERPAPER, ByVal strSavePath As String) As Boolean
    
    objPaper.PaperSize = GetSetting("ZLSOFT", strSavePath, "纸张", objPaper.PaperSize)
'    mstrPagerName = GetSetting("ZLSOFT", strSavePath, "纸张名称", "")
    objPaper.Width = GetSetting("ZLSOFT", strSavePath, "宽度", objPaper.Width)
    objPaper.Height = GetSetting("ZLSOFT", strSavePath, "高度", objPaper.Height)
    objPaper.Orientation = GetSetting("ZLSOFT", strSavePath, "纸向", objPaper.Orientation)
'    mintBin = GetSetting("ZLSOFT", strSavePath, "进纸", Printer.PaperBin)
    objPaper.BorderLeft = GetSetting("ZLSOFT", strSavePath, "左边距", objPaper.BorderLeft)
    objPaper.BorderRight = GetSetting("ZLSOFT", strSavePath, "右边距", objPaper.BorderRight)
    objPaper.BorderTop = GetSetting("ZLSOFT", strSavePath, "上边距", objPaper.BorderTop)
    objPaper.BorderBottom = GetSetting("ZLSOFT", strSavePath, "下边距", objPaper.BorderBottom)
    
    GetPaper = True
    
End Function

Public Function InitData(ByVal frmMain As Object, Optional ByVal strSavePath As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
'    mblnOK = False
'    mintPaperSize = intPaperSize
'    mblnModifyPaper = blnModifyPaper
    
    mstrSavePath = strSavePath

    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(Optional ByVal intPaperSize As Integer, Optional ByVal blnModifyPaper As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mintPaperSize = intPaperSize
    mblnModifyPaper = blnModifyPaper
    
    If ExecuteCommand("刷新数据") = False Then Exit Function
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    RefreshData = True
    
End Function


Public Function ValIDData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    ValIDData = ExecuteCommand("校验数据")
    
End Function

Public Function SaveData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    '保存打印参数
    SaveSetting "ZLSOFT", mstrSavePath, "打印机", mstrPrinter
    SaveSetting "ZLSOFT", mstrSavePath, "纸张", mintPage
    SaveSetting "ZLSOFT", mstrSavePath, "纸张名称", cboPage.Text
    SaveSetting "ZLSOFT", mstrSavePath, "宽度", mlngWidth
    SaveSetting "ZLSOFT", mstrSavePath, "高度", mlngHeight
    SaveSetting "ZLSOFT", mstrSavePath, "纸向", mintOrient
    SaveSetting "ZLSOFT", mstrSavePath, "进纸", mintBin
    SaveSetting "ZLSOFT", mstrSavePath, "左边距", mlngLeft
    SaveSetting "ZLSOFT", mstrSavePath, "右边距", mlngRight
    SaveSetting "ZLSOFT", mstrSavePath, "上边距", mlngTop
    SaveSetting "ZLSOFT", mstrSavePath, "下边距", mlngBottom
        
    SaveData = True
    
End Function

Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim strSql As String
    Dim strTmp As String
    Dim varTmp As Variant
    Dim lngTmp As Long
    Dim lngCount As Long
    Dim strPaperSize As String * 300
    Dim strPaperBin As String * 100
    Dim strPaperBinName As String * 1000
    Dim blnOK As Boolean
    Dim dblRight As Double
    Dim dblDown As Double
        
    On Error GoTo errHand
    
    mblnReading = True
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
                        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        If Printers.count = 0 Then
            MsgBox "系统中没有安装任何打印机,请先安装打印机！", vbInformation, ParamInfo.系统名称
            Unload Me: Exit Function
        End If
        
        '
        '--------------------------------------------------------------------------------------------------------------
        With cboPrinter
            .Clear
            For intLoop = 0 To Printers.count - 1
                .AddItem Printers(intLoop).DeviceName
                .ItemData(.ListCount - 1) = intLoop
            Next
        End With
        
        mstrPrinter = GetSetting("ZLSOFT", mstrSavePath, "打印机", Printer.DeviceName)
        
        On Error Resume Next
        cboPrinter.Text = mstrPrinter
'        Call gobjControl.CboLocate(cboPrinter, mstrPrinter, False)
        On Error GoTo errHand
        
        If cboPrinter.ListCount > 0 And cboPrinter.ListIndex = -1 Then
            cboPrinter.Text = Printer.DeviceName
'             Call gobjControl.CboLocate(cboPrinter, Printer.DeviceName)
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        If mblnModifyPaper = False Then
            cboPage.Enabled = False
            UD上.Enabled = False
            UD下.Enabled = False
            UD左.Enabled = False
            UD右.Enabled = False
            
            txt上.Enabled = False
            txt下.Enabled = False
            txt左.Enabled = False
            txt右.Enabled = False
            
            opt纵向.Enabled = False
            opt横向.Enabled = False
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "装载打印机相关信息"
        
        
        If cboPrinter.ListIndex = -1 Then Exit Function
        
        Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
        mstrPrinter = Printer.DeviceName
        lblLoc.Caption = "位置: " & Printer.Port
    
        '如果支持,则保持原有纸张
        If mintPage <> 256 Then
            On Error Resume Next
            Printer.PaperSize = mintPage
            On Error GoTo errHand
            mintPage = Printer.PaperSize
        End If
    
        
        
        '装载纸张大小
        '--------------------------------------------------------------------------------------------------------------
        cboPage.Clear
        lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaperSize, 0)
        For intLoop = 1 To lngCount
            lngTmp = Asc(Mid(strPaperSize, intLoop * 2, 1)) * 256 + Asc(Mid(strPaperSize, intLoop * 2 - 1, 1))
    
            If mbytMode = 1 Then
                If lngTmp = 9 Or lngTmp = 13 Then
                    cboPage.AddItem mclsPrint.GetPaperName(lngTmp)
                    cboPage.ItemData(cboPage.ListCount - 1) = lngTmp
                    If lngTmp = mintPage Then cboPage.ListIndex = cboPage.NewIndex
                End If
            Else
                If lngTmp >= 1 And lngTmp <= 41 Then '只列出标准支持的纸张
                    cboPage.AddItem mclsPrint.GetPaperName(lngTmp)
                    cboPage.ItemData(cboPage.ListCount - 1) = lngTmp
                    If lngTmp = mintPage Then cboPage.ListIndex = cboPage.NewIndex
                End If
            End If
        Next

        
        '自定义纸张处理
        '--------------------------------------------------------------------------------------------------------------
        lngTmp = 256
        cboPage.AddItem mclsPrint.GetPaperName(lngTmp)
        cboPage.ItemData(cboPage.ListCount - 1) = lngTmp
        If mintPage = 256 Then cboPage.ListIndex = cboPage.NewIndex
        If cboPage.ListIndex = -1 And cboPage.ListCount > 0 Then cboPage.ListIndex = 0

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '如果支持,则保持原有进纸方式
        On Error Resume Next
        Printer.PaperBin = mintBin
        On Error GoTo errHand
        mintBin = Printer.PaperBin
    
        '设置可用进纸方式
        cboBin.Clear
        '--------------------------------------------------------------------------------------------------------------
        lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
        lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
        lngTmp = 1
        For intLoop = 1 To lngCount
            '进纸名称
            Do
                If Mid(strPaperBinName, lngTmp, 1) = Chr(0) Then
                    cboBin.AddItem Trim(strTmp)
    
                    '进纸编号
                    cboBin.ItemData(cboBin.ListCount - 1) = Asc(Mid(strPaperBin, intLoop * 2, 1)) * 256 + Asc(Mid(strPaperBin, intLoop * 2 - 1, 1))
                    If cboBin.ItemData(cboBin.ListCount - 1) = mintBin Then
                        cboBin.ListIndex = cboBin.NewIndex
                    End If
    
                    lngTmp = 24 + lngTmp - LenB(StrConv(strTmp, vbFromUnicode))
                    strTmp = ""
                    Exit Do
                Else
                    strTmp = strTmp & Mid(strPaperBinName, lngTmp, 1)
                    lngTmp = lngTmp + 1
                End If
            Loop
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case "装载纸张信息"

        Select Case cboPage.ItemData(cboPage.ListIndex)
        Case 256
            '强行设置自定义纸张可用,不检查
            mintPage = 256
        Case Else
            Printer.PaperSize = cboPage.ItemData(cboPage.ListIndex)
            mintPage = Printer.PaperSize
        End Select
    
        opt纵向.Enabled = True
        opt横向.Enabled = True
                
        '--------------------------------------------------------------------------------------------------------------
        err = 0
        On Error Resume Next
        Printer.Orientation = 1
        If Printer.Orientation = 1 Then
            Printer.Orientation = 2
            If Printer.Orientation <> 2 Then
                opt纵向.Enabled = False
                opt横向.Enabled = False
            End If
        End If
        
'        opt纵向.Enabled = mblnModifyPaper
'        opt横向.Enabled = mblnModifyPaper

        On Error GoTo errHand
        
        '--------------------------------------------------------------------------------------------------------------
        If opt横向.Enabled Then
            Printer.Orientation = mintOrient
            mintOrient = Printer.Orientation
        Else
            opt纵向.Value = True
            img纵向.Visible = True
            img横向.Visible = False
        End If
    
        '最后实际设置纸张大小(纸向影响之后)
        Select Case mintPage
        Case 256
            '自定义纸张认为全部可以打印
            mdblW = 0
            mdblH = 0
    
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            UDWidth.Enabled = True
            UDHeight.Enabled = True
        Case Else
            '取该打印机支持该幅面的真实尺寸
            mlngWidth = Printer.Width
            mlngHeight = Printer.Height
    
            '不可打印区域比例
            mdblW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWidth)
            mdblH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    
            txtWidth.Enabled = False
            txtHeight.Enabled = False
            UDWidth.Enabled = False
            UDHeight.Enabled = False
    
        End Select
    
        '显示纸张尺寸
        txtWidth.Tag = mlngWidth
        txtWidth.Text = CLng(mlngWidth / 56.7)
        txtHeight.Tag = mlngHeight
        txtHeight.Text = CLng(mlngHeight / 56.7)
    
        '显示可用边距
        '最小在可打印区域之内
        '最大不超过宽高的1/4
    '    If cboPage.Text = "B5, 182 x 257 毫米" Then
    '        UD左.Min = 0
    '        UD左.Max = 5
    '    Else
            UD左.Min = mlngWidth / 56.7 * mdblW
            UD左.Max = mlngWidth / 56.7 / 4
    '    End If
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
        If mintOrient = 1 Then
            opt纵向.Value = True: opt纵向_Click
        Else
            opt横向.Value = True: opt横向_Click
        End If
            
        '边距
        txt左.Text = mlngLeft
        txt右.Text = mlngRight
        txt上.Text = mlngTop
        txt下.Text = mlngBottom
        
    '--------------------------------------------------------------------------------------------------------------
    Case "校验数据"
            
        If Not IsNumeric(txtWidth.Text) Then
            MsgBox "请确定报表的纸张宽度！", vbExclamation, App.Title
            txtWidth.SetFocus: Exit Function
        End If
        If CInt(txtWidth.Text) > UDWidth.Max Then
            MsgBox "报表的纸张宽度不能超过" & UDWidth.Max & "毫米！", vbExclamation, App.Title
            txtWidth.SetFocus: Exit Function
        End If
    
        If Not IsNumeric(txtHeight.Text) Then
            MsgBox "请确定报表的纸张高度！", vbExclamation, App.Title
            txtWidth.SetFocus: Exit Function
        End If
        If CInt(txtHeight.Text) > UDHeight.Max Then
            MsgBox "报表的纸张高度不能超过" & UDHeight.Max & "毫米！", vbExclamation, App.Title
            txtHeight.SetFocus: Exit Function
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "显示纸张样式"
    
        On Error Resume Next
    
        picPaper.cls
    
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
        On Error GoTo errHand
        
    '--------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        mintPage = GetSetting("ZLSOFT", mstrSavePath, "纸张", Printer.PaperSize)
        mstrPagerName = GetSetting("ZLSOFT", mstrSavePath, "纸张名称", "")
        mlngWidth = GetSetting("ZLSOFT", mstrSavePath, "宽度", Printer.Width)
        mlngHeight = GetSetting("ZLSOFT", mstrSavePath, "高度", Printer.Height)
        mintOrient = GetSetting("ZLSOFT", mstrSavePath, "纸向", Printer.Orientation)
        mintBin = GetSetting("ZLSOFT", mstrSavePath, "进纸", Printer.PaperBin)
        mlngLeft = GetSetting("ZLSOFT", mstrSavePath, "左边距", OFFSET_LEFT)
        mlngRight = GetSetting("ZLSOFT", mstrSavePath, "右边距", OFFSET_RIGHT)
        mlngTop = GetSetting("ZLSOFT", mstrSavePath, "上边距", OFFSET_TOP)
        mlngBottom = GetSetting("ZLSOFT", mstrSavePath, "下边距", OFFSET_BOTTOM)
        mlngLeft = GetSetting("ZLSOFT", mstrSavePath, "左边距", OFFSET_LEFT)
        mlngRight = GetSetting("ZLSOFT", mstrSavePath, "右边距", OFFSET_RIGHT)
        mlngTop = GetSetting("ZLSOFT", mstrSavePath, "上边距", OFFSET_TOP)
        mlngBottom = GetSetting("ZLSOFT", mstrSavePath, "下边距", OFFSET_BOTTOM)
        
        Call ExecuteCommand("装载打印机相关信息")
        Call ExecuteCommand("装载纸张信息")
        Call ExecuteCommand("显示纸张样式")
        
    End Select

    ExecuteCommand = True
    
    mblnReading = False
    
    Exit Function
    
    '
    '----------------------------------------------------------------------------------------------------------
errHand:

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog

End Function

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

'######################################################################################################################

Private Sub cboBin_Click()
    If cboBin.ListIndex <> -1 Then
        mintBin = cboBin.ItemData(cboBin.ListIndex)
    End If
End Sub

Private Sub cboPage_Click()
    
    If mblnReading Then Exit Sub
    
    Call ExecuteCommand("装载纸张信息")
    Call ExecuteCommand("显示纸张样式")
    
End Sub

Private Sub cboPrinter_Click()

    If mblnReading Then Exit Sub
    
    Call ExecuteCommand("装载打印机相关信息")
    Call ExecuteCommand("装载纸张信息")
    Call ExecuteCommand("显示纸张样式")
    Call ExecuteCommand("控件状态")
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
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

        Call cboPage_Click
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

        Call cboPage_Click
    End If
End Sub

Private Sub txtHeight_Change()
    
    If mblnReading Then Exit Sub
    
    If IsNumeric(txtHeight.Text) Then
        txtHeight.Tag = CLng(txtHeight.Text * 56.7)
        mlngHeight = CLng(txtHeight.Text * 56.7)

        cboPage.ListIndex = cboPage.ListCount - 1
    End If
    
    Call ExecuteCommand("显示纸张样式")
End Sub

Private Sub txtWidth_Change()
    
    If mblnReading Then Exit Sub
    
    If IsNumeric(txtWidth.Text) Then
        txtWidth.Tag = CLng(txtWidth.Text * 56.7)
        mlngWidth = CLng(txtWidth.Text * 56.7)

        cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ExecuteCommand("显示纸张样式")
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
    gobjControl.TxtSelAll txt上
End Sub

Private Sub txt上_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt下_GotFocus()
    gobjControl.TxtSelAll txt下
End Sub

Private Sub txt下_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt右_GotFocus()
    gobjControl.TxtSelAll txt右
End Sub

Private Sub txt左_GotFocus()
    gobjControl.TxtSelAll txt左
End Sub

Private Sub txt左_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub UD上_Change()
    If mblnReading Then Exit Sub
    mlngTop = UD上.Value
    Call ExecuteCommand("显示纸张样式")
End Sub

Private Sub UD下_Change()
    If mblnReading Then Exit Sub
    mlngBottom = UD下.Value
    Call ExecuteCommand("显示纸张样式")
End Sub

Private Sub UD右_Change()
    If mblnReading Then Exit Sub
    mlngRight = UD右.Value
    Call ExecuteCommand("显示纸张样式")
End Sub

Private Sub UD左_Change()
    If mblnReading Then Exit Sub
    mlngLeft = UD左.Value
    Call ExecuteCommand("显示纸张样式")
End Sub

