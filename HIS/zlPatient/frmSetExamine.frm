VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetExamine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置参数"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "frmSetExamine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&O)"
      Height          =   350
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpE 
      Height          =   300
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   74973187
      CurrentDate     =   37068
   End
   Begin MSComCtl2.DTPicker dtpB 
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   74973187
      CurrentDate     =   37068
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出院时间"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   2497
      TabIndex        =   2
      Top             =   420
      Width           =   180
   End
End
Attribute VB_Name = "frmSetExamine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjForm As Object

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
     mobjForm.mdtBegin = dtpB.Value
     mobjForm.mdtEnd = dtpE.Value
     Unload Me
End Sub

Public Sub EditWhere(frmMain As Object)
    Set mobjForm = frmMain
    dtpB.Value = mobjForm.mdtBegin
    dtpE.Value = mobjForm.mdtEnd
    frmSetExamine.Show 1, frmMain
End Sub
