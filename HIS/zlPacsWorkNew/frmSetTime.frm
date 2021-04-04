VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置时间"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   Icon            =   "frmSetTime.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确定"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   570
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   145686531
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   1290
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   145686531
      CurrentDate     =   38082.9993055556
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      Caption         =   "结束时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblBegin 
      AutoSize        =   -1  'True
      Caption         =   "开始时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1200
   End
End
Attribute VB_Name = "frmSetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mdtBegin As Date
Private mdtEnd As Date

Public Function ShowSetTime(ByRef dtBegin As Date, ByRef dtEnd As Date, ower As Object) As Boolean
    mblnOk = False
    mdtBegin = dtBegin
    mdtEnd = dtEnd
    
    Me.Show 1, ower
    
    dtBegin = mdtBegin
    dtEnd = mdtEnd
    ShowSetTime = mblnOk
End Function

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mblnOk = False
    
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "提示"
    Err.Clear
End Sub

Private Sub cmdSure_Click()
    On Error GoTo errHandle
    
    mblnOk = True
    mdtBegin = dtpBegin.Value
    mdtEnd = dtpEnd.Value
    
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "提示"
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    dtpBegin.Value = mdtBegin
    dtpEnd.Value = mdtEnd
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "提示"
    Err.Clear
End Sub



