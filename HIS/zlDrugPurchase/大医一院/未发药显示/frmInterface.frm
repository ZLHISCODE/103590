VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "显示界面调整"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboScreen 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   3435
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   8040
      TabIndex        =   40
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用(&A)"
      Height          =   375
      Left            =   6840
      TabIndex        =   39
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefualt 
      Caption         =   "缺省值(&D)"
      Height          =   375
      Left            =   5640
      TabIndex        =   38
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame fraView 
      Caption         =   "待取药病人姓名显示"
      Height          =   1215
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   9135
      Begin VB.TextBox txtSeconds 
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   720
         Width           =   585
      End
      Begin VB.TextBox txtColumns 
         Height          =   270
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   585
      End
      Begin VB.ComboBox cboBackColor 
         Height          =   300
         Index           =   2
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboForeColor 
         Height          =   300
         Index           =   2
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtSize 
         Height          =   270
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   585
      End
      Begin VB.ComboBox cboFont 
         Height          =   300
         Index           =   2
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin ComCtl2.UpDown udSize 
         Height          =   270
         Index           =   2
         Left            =   3600
         TabIndex        =   25
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         Value           =   36
         BuddyControl    =   "txtSize(2)"
         BuddyDispid     =   196618
         BuddyIndex      =   2
         OrigLeft        =   3600
         OrigTop         =   360
         OrigRight       =   3855
         OrigBottom      =   630
         Max             =   100
         Min             =   12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udColumns 
         Height          =   270
         Left            =   3600
         TabIndex        =   35
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         Value           =   4
         BuddyControl    =   "txtColumns"
         BuddyDispid     =   196615
         OrigLeft        =   3600
         OrigTop         =   720
         OrigRight       =   3855
         OrigBottom      =   990
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udSeconds 
         Height          =   270
         Left            =   1800
         TabIndex        =   32
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         Value           =   30
         BuddyControl    =   "txtSeconds"
         BuddyDispid     =   196614
         OrigLeft        =   1440
         OrigTop         =   720
         OrigRight       =   1695
         OrigBottom      =   990
         Increment       =   10
         Max             =   300
         Min             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblSeconds 
         AutoSize        =   -1  'True
         Caption         =   "刷新秒数"
         Height          =   180
         Left            =   360
         TabIndex        =   30
         Top             =   770
         Width           =   720
      End
      Begin VB.Label lblColumns 
         AutoSize        =   -1  'True
         Caption         =   "列数"
         Height          =   180
         Left            =   2520
         TabIndex        =   33
         Top             =   770
         Width           =   360
      End
      Begin VB.Label lblBackColor 
         AutoSize        =   -1  'True
         Caption         =   "字体背景"
         Height          =   180
         Index           =   2
         Left            =   6840
         TabIndex        =   28
         Top             =   410
         Width           =   720
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         Caption         =   "字体颜色"
         Height          =   180
         Index           =   2
         Left            =   4320
         TabIndex        =   26
         Top             =   410
         Width           =   720
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "字号"
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   23
         Top             =   410
         Width           =   360
      End
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         Caption         =   "字体"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   21
         Top             =   410
         Width           =   360
      End
   End
   Begin VB.Frame fraView 
      Caption         =   "待取药病人显示"
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   9135
      Begin VB.ComboBox cboFont 
         Height          =   300
         Index           =   1
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtSize 
         Height          =   270
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   585
      End
      Begin VB.ComboBox cboForeColor 
         Height          =   300
         Index           =   1
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboBackColor 
         Height          =   300
         Index           =   1
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin ComCtl2.UpDown udSize 
         Height          =   270
         Index           =   1
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         Value           =   36
         BuddyControl    =   "txtSize(1)"
         BuddyDispid     =   196618
         BuddyIndex      =   1
         OrigLeft        =   3600
         OrigTop         =   360
         OrigRight       =   3855
         OrigBottom      =   630
         Max             =   100
         Min             =   12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         Caption         =   "字体"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   410
         Width           =   360
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "字号"
         Height          =   180
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   410
         Width           =   360
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         Caption         =   "字体颜色"
         Height          =   180
         Index           =   1
         Left            =   4320
         TabIndex        =   16
         Top             =   410
         Width           =   720
      End
      Begin VB.Label lblBackColor 
         AutoSize        =   -1  'True
         Caption         =   "字体背景"
         Height          =   180
         Index           =   1
         Left            =   6840
         TabIndex        =   18
         Top             =   410
         Width           =   720
      End
   End
   Begin VB.Frame fraView 
      Caption         =   "发药窗口号显示"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.ComboBox cboBackColor 
         Height          =   300
         Index           =   0
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboForeColor 
         Height          =   300
         Index           =   0
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin ComCtl2.UpDown udSize 
         Height          =   270
         Index           =   0
         Left            =   3586
         TabIndex        =   5
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   327681
         Value           =   36
         BuddyControl    =   "txtSize(0)"
         BuddyDispid     =   196618
         BuddyIndex      =   0
         OrigLeft        =   3600
         OrigTop         =   360
         OrigRight       =   3855
         OrigBottom      =   630
         Max             =   100
         Min             =   12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSize 
         Height          =   270
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   585
      End
      Begin VB.ComboBox cboFont 
         Height          =   300
         Index           =   0
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblBackColor 
         AutoSize        =   -1  'True
         Caption         =   "字体背景"
         Height          =   180
         Index           =   0
         Left            =   6840
         TabIndex        =   8
         Top             =   410
         Width           =   720
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         Caption         =   "字体颜色"
         Height          =   180
         Index           =   0
         Left            =   4320
         TabIndex        =   6
         Top             =   410
         Width           =   720
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "字号"
         Height          =   180
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   410
         Width           =   360
      End
      Begin VB.Label lblFont 
         AutoSize        =   -1  'True
         Caption         =   "字体"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   410
         Width           =   360
      End
   End
   Begin VB.Label lblScreen 
      AutoSize        =   -1  'True
      Caption         =   "显示屏幕"
      Height          =   180
      Left            =   240
      TabIndex        =   36
      Top             =   3480
      Width           =   720
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdApply_Click()
    On Error Resume Next
    SaveSetting "ZLSOFT", "未发药病人显示\发药窗口显示", "字体", cboFont(0).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\待取药病人显示", "字体", cboFont(1).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\病人姓名显示", "字体", cboFont(2).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\发药窗口显示", "字号", udSize(0).Value
    SaveSetting "ZLSOFT", "未发药病人显示\待取药病人显示", "字号", udSize(1).Value
    SaveSetting "ZLSOFT", "未发药病人显示\病人姓名显示", "字号", udSize(2).Value
    SaveSetting "ZLSOFT", "未发药病人显示\发药窗口显示", "字体前景色", cboForeColor(0).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\待取药病人显示", "字体前景色", cboForeColor(1).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\病人姓名显示", "字体前景色", cboForeColor(2).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\发药窗口显示", "字体背景色", cboBackColor(0).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\待取药病人显示", "字体背景色", cboBackColor(1).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\病人姓名显示", "字体背景色", cboBackColor(2).ListIndex
    SaveSetting "ZLSOFT", "未发药病人显示\病人姓名显示", "刷新秒数", udSeconds.Value
    SaveSetting "ZLSOFT", "未发药病人显示\病人姓名显示", "列数", udColumns.Value
    SaveSetting "ZLSOFT", "未发药病人显示", "显示屏幕", cboScreen.ListIndex
    Unload Me
    Call SetWindow
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefualt_Click()
    Call SetInterface(True)
End Sub

Private Sub Form_Load()
    Call SetInterface
End Sub

Private Sub SetInterface(Optional ByVal blnDefault As Boolean = False)
    Dim i As Byte
    Dim strVal As String
    
    If Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="刷新秒数", Default:="")) = 0 Then
        blnDefault = True
    End If
    
    '屏幕
    With cboScreen
        .Clear
        .AddItem "主屏"
        If GetSystemMetrics(80) > 1 Then
            .AddItem "次屏"
            .ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示", Key:="显示屏幕", Default:=""))
        Else
            .ListIndex = 0
        End If
    End With
    '字体
    For i = 0 To 2
        With cboFont(i)
            .Clear
            .AddItem "宋体"
            .AddItem "黑体"
            .AddItem "楷体_GB2312"
            .AddItem "楷体"
        End With
    Next
    If blnDefault Then
        cboFont(0).ListIndex = 0
        cboFont(1).ListIndex = 0
        cboFont(2).ListIndex = 0
    Else
        cboFont(0).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\发药窗口显示", Key:="字体", Default:=""))
        cboFont(1).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\待取药病人显示", Key:="字体", Default:=""))
        cboFont(2).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="字体", Default:=""))
    End If
    '字号
    With udSize(0)
        If blnDefault Then
            strVal = "60"
        Else
            strVal = GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\发药窗口显示", Key:="字号", Default:="")
        End If
        .Max = 100: .Min = 12: .Value = IIf(Trim(strVal) = "", .Min, strVal)
    End With
    With udSize(1)
        If blnDefault Then
            strVal = "60"
        Else
            strVal = GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\待取药病人显示", Key:="字号", Default:="")
        End If
        .Max = 100: .Min = 12: .Value = IIf(Trim(strVal) = "", .Min, strVal)
    End With
    With udSize(2)
        If blnDefault Then
            strVal = "72"
        Else
            strVal = GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="字号", Default:="")
        End If
        .Max = 100: .Min = 12: .Value = IIf(Trim(strVal) = "", .Min, strVal)
    End With
    '颜色
    For i = 0 To 2
        With cboForeColor(i)
            .Clear
            .AddItem "白色": .ItemData(.NewIndex) = vbWhite
            .AddItem "红色": .ItemData(.NewIndex) = vbRed
            .AddItem "蓝色": .ItemData(.NewIndex) = vbBlue
            .AddItem "黄色": .ItemData(.NewIndex) = vbYellow
            .AddItem "绿色": .ItemData(.NewIndex) = vbGreen
            .AddItem "黑色": .ItemData(.NewIndex) = vbBlack
        End With
    Next
    For i = 0 To 2
        With cboBackColor(i)
            .Clear
            .AddItem "蓝色": .ItemData(.NewIndex) = vbBlue
            .AddItem "红色": .ItemData(.NewIndex) = vbRed
            .AddItem "黄色": .ItemData(.NewIndex) = vbYellow
            .AddItem "绿色": .ItemData(.NewIndex) = vbGreen
            .AddItem "白色": .ItemData(.NewIndex) = vbWhite
            .AddItem "黑色": .ItemData(.NewIndex) = vbBlack
        End With
    Next
    If blnDefault Then
        cboForeColor(0).ListIndex = 0
        cboForeColor(1).ListIndex = 0
        cboForeColor(2).ListIndex = 2
        cboBackColor(0).ListIndex = 0
        cboBackColor(1).ListIndex = 0
        cboBackColor(2).ListIndex = 0
    Else
        cboForeColor(0).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\发药窗口显示", Key:="字体前景色", Default:=""))
        cboForeColor(1).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\待取药病人显示", Key:="字体前景色", Default:=""))
        cboForeColor(2).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="字体前景色", Default:=""))
        cboBackColor(0).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\发药窗口显示", Key:="字体背景色", Default:=""))
        cboBackColor(1).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\待取药病人显示", Key:="字体背景色", Default:=""))
        cboBackColor(2).ListIndex = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="字体背景色", Default:=""))
    End If
    '刷新秒数
    With udSeconds
        strVal = GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="刷新秒数", Default:="")
        .Max = 300: .Min = 30: .Value = IIf(Trim(strVal) = "", .Min, strVal)
    End With
    '列数
    With udColumns
        strVal = GetSetting(appName:="ZLSOFT", Section:="未发药病人显示\病人姓名显示", Key:="列数", Default:="")
        .Max = 10: .Min = 2: .Value = IIf(Trim(strVal) = "", 4, strVal)
    End With
    
End Sub

Private Sub SetWindow()
    With frmUnSendDrug
        .mstrFormFont = Me.cboFont(0).Text
        .mstrDrugFont = Me.cboFont(1).Text
        .mstrPatientFont = Me.cboFont(2).Text
        .mintFormSize = Me.udSize(0).Value
        .mintDrugSize = Me.udSize(1).Value
        .mintPatientSize = Me.udSize(2).Value
        .mlngBackColorA = Me.cboBackColor(0).ListIndex
        .mlngBackColorB = Me.cboBackColor(1).ListIndex
        .mlngBackColorC = Me.cboBackColor(2).ListIndex
        .mlngForeColorA = Me.cboForeColor(0).ListIndex
        .mlngForeColorB = Me.cboForeColor(1).ListIndex
        .mlngForeColorC = Me.cboForeColor(2).ListIndex
        .mintCols = Me.udColumns.Value
        .mbytScreen = Me.cboScreen.ListIndex
        .Tag = ""
        
        With frmUnSendDrug
            .Entry .mlngStockID, .lblFormNO.Tag
        End With
    End With
    
End Sub

