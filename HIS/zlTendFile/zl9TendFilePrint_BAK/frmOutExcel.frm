VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutExcel 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame FraBack 
      Height          =   1215
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   5595
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   210
         TabIndex        =   4
         Top             =   810
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   165
         Left            =   3090
         TabIndex        =   3
         Top             =   330
         Width           =   105
      End
      Begin VB.Label lblnum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   165
         Left            =   2790
         TabIndex        =   2
         Top             =   330
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "正在输出到EXCEL文件，已完成了"
         Height          =   165
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmOutExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'本窗体用于输出到EXCEL

Option Explicit
Dim zlExcelcls As zlExcel

Private Sub Form_Activate()
    Set zlExcelcls = New zlExcel
    Set zlExcelcls.frmTempExcel = Me
    zlExcelcls.zlExcelFile
    Set zlExcelcls = Nothing
    Unload Me
End Sub
