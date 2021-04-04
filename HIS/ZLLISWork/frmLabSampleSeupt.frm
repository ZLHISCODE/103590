VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmLabSampleSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8595
   Icon            =   "frmLabSampleSeupt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   6810
      TabIndex        =   10
      Top             =   3630
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   5220
      TabIndex        =   9
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "操作"
      Height          =   1365
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   3525
      Begin VB.CheckBox chkFindMove 
         Caption         =   "查找到病人后焦点移动到条码输入"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   3135
      End
      Begin VB.CheckBox ChkContinuous 
         Caption         =   "连续输入条形码"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   2595
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "可使用的试管"
      Height          =   3375
      Left            =   3630
      TabIndex        =   1
      Top             =   90
      Width           =   4905
      Begin XtremeReportControl.ReportControl rptCuvette 
         Height          =   2985
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4665
         _Version        =   589884
         _ExtentX        =   8229
         _ExtentY        =   5265
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         SkipGroupsFocus =   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "条码"
      Height          =   1905
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   3525
      Begin VB.CheckBox chkBackBill 
         Caption         =   "已完成采集后打印回执单"
         Height          =   225
         Left            =   300
         TabIndex        =   4
         Top             =   1440
         Width           =   2325
      End
      Begin VB.CheckBox chkComPlete 
         Caption         =   "生成或绑定条码后标志为已采集"
         Height          =   225
         Left            =   300
         TabIndex        =   3
         Top             =   900
         Width           =   2835
      End
      Begin VB.CheckBox ChkBarCodePrint 
         Caption         =   "生成或绑定条码后打印条码"
         Height          =   225
         Left            =   300
         TabIndex        =   2
         Top             =   390
         Width           =   2715
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   120
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":0078
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":0612
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":0BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabSampleSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCuvette                               '试管
    选择
    编码
    名称
    添加剂
    采血量
    规格
    颜色
End Enum

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim Record As ReportRecord
    Dim Column As ReportColumn
    
    rptCuvette.SetImageList ImgList
    With Me.rptCuvette.Columns
        Set Column = .Add(mCuvette.选择, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mCuvette.编码, "编码", 55, True)
        Set Column = .Add(mCuvette.名称, "名称", 80, True)
        Set Column = .Add(mCuvette.添加剂, "添加剂", 90, True)
        Set Column = .Add(mCuvette.采血量, "采血量", 60, True)
        Set Column = .Add(mCuvette.规格, "规格", 60, True)
        Set Column = .Add(mCuvette.颜色, "", 18, True): Column.Icon = 3
    End With
    
    gstrSql = "Select 编码,名称,简码,添加剂,采血量,规格,颜色 From 采血管类型"
    zlDatabase.OpenRecordset rsTmp, gstrSql, gstrSysName
    Do While Not rsTmp.EOF
        Set Record = Me.rptCuvette.Records.Add
        For intLoop = 0 To Me.rptCuvette.Columns.Count
            Record.AddItem ""
        Next
        
        Record(mCuvette.选择).HasCheckbox = True
        Record(mCuvette.选择).Checked = True
        Record(mCuvette.编码).Value = Nvl(rsTmp("编码"))
        Record(mCuvette.名称).Value = Nvl(rsTmp("名称"))
        Record(mCuvette.添加剂).Value = Nvl(rsTmp("添加剂"))
        Record(mCuvette.采血量).Value = Nvl(rsTmp("采血量"))
        Record(mCuvette.规格).Value = Nvl(rsTmp("规格"))
        Record(mCuvette.颜色).BackColor = Nvl(rsTmp("颜色"))
        
        For intLoop = 0 To Me.rptCuvette.Columns.Count
            Record(intLoop).ForeColor = Nvl(rsTmp("颜色"))
        Next
        
        rsTmp.MoveNext
    Loop
    Me.rptCuvette.Populate
End Sub
