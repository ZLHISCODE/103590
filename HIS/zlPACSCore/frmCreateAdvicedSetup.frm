VERSION 5.00
Begin VB.Form frmCreateAdvicedSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "刻录高级设置"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "frmCreateAdvicedSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdDefault 
      Caption         =   "默认(&D)"
      Height          =   350
      Left            =   120
      TabIndex        =   12
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6450
      TabIndex        =   11
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4740
      TabIndex        =   10
      Top             =   2460
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "写入选项"
      Height          =   2265
      Left            =   4260
      TabIndex        =   1
      Top             =   60
      Width           =   3915
      Begin VB.CheckBox ChkWriterAutoVerify 
         Caption         =   "自动校验文件数据"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   1710
         Width           =   3195
      End
      Begin VB.CheckBox ChkWriterBufferProof 
         Caption         =   "缓冲区校验"
         CausesValidation=   0   'False
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   1290
         Width           =   2235
      End
      Begin VB.CheckBox ChkWriterTestWriter 
         Caption         =   "测试写入(DVD格式无效)"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   3195
      End
      Begin VB.CheckBox ChkWriterCloseDisk 
         Caption         =   "结束光盘(不可再写入)"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   2805
      End
      Begin VB.CheckBox ChkWriterCheckImage 
         Caption         =   "不使用高速缓存写入(限CD_RW)"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   3525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据选项"
      Height          =   2265
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4065
      Begin VB.CheckBox ChkDataHighCompatibilityMode 
         Caption         =   "高兼容性DVD(写入文件最小要达到1GB)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   990
         Width           =   3585
      End
      Begin VB.CheckBox ChkDataCDRWMode 
         Caption         =   "CDR/W写入时使用模式"
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   630
         Width           =   3555
      End
      Begin VB.CheckBox ChkDataUseJoliet 
         Caption         =   "使用Joliet(使文件名最大可达64个字符)"
         Height          =   345
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmCreateAdvicedSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDefault_Click()
    '数据
    ChkDataUseJoliet.Value = 1
    ChkDataCDRWMode.Value = 0
    ChkDataHighCompatibilityMode.Value = 0
    '写入
    ChkWriterCheckImage.Value = 0
    ChkWriterCloseDisk.Value = 1
    ChkWriterTestWriter.Value = 1
    ChkWriterBufferProof.Value = 1
    ChkWriterAutoVerify.Value = 1
End Sub

Private Sub cmdOK_Click()
    SaveOrLoadSetup 1
    Unload Me
End Sub



Private Sub Form_Load()
    SaveOrLoadSetup 2
End Sub

Sub SaveOrLoadSetup(SaveOrLoad As Integer)
    '读入或保存参数
    'SaveOrLoad = 1 保存 = 2 读入
    
    Dim intUseJoliet As Integer
    Dim intCDRWMode  As Integer
    Dim blHighCompatibilityMode  As Boolean
    Dim blCheckImage As Boolean
    Dim blCloseDisk As Boolean
    Dim blTestWriter As Boolean
    Dim blBufferProof As Boolean
    Dim blAutoVerify As Boolean
    
    '保存
    If SaveOrLoad = 1 Then
        intUseJoliet = IIf((ChkDataUseJoliet.Value = vbChecked), vtyISO9660_JOLIET, vtyISO9660_ONLY)
        intCDRWMode = IIf((ChkDataCDRWMode.Value = vbChecked), wtpDataMode2_XA, wtpDataMode1)
        blHighCompatibilityMode = (ChkDataHighCompatibilityMode.Value = vbChecked)
        blCheckImage = (ChkWriterCheckImage.Value = vbChecked)
        blCloseDisk = (ChkWriterCloseDisk.Value = vbChecked)
        blTestWriter = (ChkWriterTestWriter.Value = vbChecked)
        blBufferProof = (ChkWriterBufferProof.Value = vbChecked)
        blAutoVerify = (ChkWriterAutoVerify.Value = vbChecked)
    
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "使用Joliet", intUseJoliet
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "使用CDRW模式", intCDRWMode
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "高兼容DVD模式", blHighCompatibilityMode
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "不使用高速缓存", blCheckImage
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "关闭光盘", blCloseDisk
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "测试写入", blTestWriter
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "缓存校验", blBufferProof
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "自动数据校验", blAutoVerify
    Else
        ChkDataUseJoliet.Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "使用Joliet", 1)
        intCDRWMode = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "使用CDRW模式", 1)
        ChkDataCDRWMode.Value = IIf(intCDRWMode = 2, 1, 0)
        blHighCompatibilityMode = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "高兼容DVD模式", 0)
        ChkDataHighCompatibilityMode.Value = IIf(blHighCompatibilityMode, 1, 0)
        blCheckImage = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "不使用高速缓存", 0)
        ChkWriterCheckImage.Value = IIf(blCheckImage, 1, 0)
        blCloseDisk = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "关闭光盘", 1)
        ChkWriterCloseDisk.Value = IIf(blCloseDisk, 1, 0)
        blTestWriter = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "测试写入", 1)
        ChkWriterTestWriter.Value = IIf(blTestWriter, 1, 0)
        blBufferProof = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "缓存校验", 1)
        ChkWriterBufferProof.Value = IIf(blBufferProof, 1, 0)
        blAutoVerify = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\刻录设置", "自动数据校验", 1)
        ChkWriterAutoVerify.Value = IIf(blAutoVerify, 1, 0)
    End If
    
End Sub
