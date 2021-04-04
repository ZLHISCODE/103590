VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPubResource 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfPrint 
      Height          =   465
      Left            =   1770
      TabIndex        =   0
      Top             =   75
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.TaskPanel tpl 
      Height          =   390
      Left            =   495
      TabIndex        =   1
      Top             =   45
      Width           =   1110
      _Version        =   589884
      _ExtentX        =   1958
      _ExtentY        =   688
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2475
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":0000
            Key             =   "µ±«∞"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":6862
            Key             =   "÷∏ æ∆˜"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":D0C4
            Key             =   "Õº±Í"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":D45E
            Key             =   "—°‘Ò"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":DE70
            Key             =   "æ‹æ¯…Í«Î"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":E40A
            Key             =   "ΩË‘ƒ…Í«Î"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":14C6C
            Key             =   "≈˙◊º…Í«Î"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1B4CE
            Key             =   "refusepatient"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1BA68
            Key             =   "waitpatient"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1C002
            Key             =   "questionpatient2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1C59C
            Key             =   "aduitpatient2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1CFAE
            Key             =   "aduitpatient"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1D9C0
            Key             =   "dealpatient2"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1E3D2
            Key             =   "questionpatient"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1EDE4
            Key             =   "dealpatient"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":1F7F6
            Key             =   "patient"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":26058
            Key             =   "lockpatient"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2C8BA
            Key             =   "archivepatient"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2CE54
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2D3EE
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2D988
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2DF22
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2E4BC
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2EA56
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":2EFF0
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":35852
            Key             =   "person"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":3C0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":3C64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":3CBE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   3315
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3915
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   0
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPubResource.frx":3CCCC
   End
   Begin XtremeCommandBars.ImageManager imgApp 
      Left            =   1185
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPubResource.frx":3F1B8
   End
End
Attribute VB_Name = "frmPubResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

