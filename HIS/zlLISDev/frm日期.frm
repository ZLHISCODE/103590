VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm日期 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择导入数据的日期"
   ClientHeight    =   1695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3600
   Icon            =   "frm日期.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.DTPicker dtp日期 
      Height          =   480
      Left            =   720
      TabIndex        =   2
      Top             =   300
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   3735555
      CurrentDate     =   40238
      MaxDate         =   401769
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2115
      TabIndex        =   1
      Top             =   1095
      Width           =   1100
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   360
      TabIndex        =   0
      Top             =   1095
      Width           =   1100
   End
End
Attribute VB_Name = "frm日期"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mblnOK As Boolean
Private mdateIn As Date

Public Function ShowMe(ByRef str日期 As String) As Boolean
    mblnOK = False
    mdateIn = CDate(str日期)
    Me.Show vbModal
    If mblnOK = True Then
        str日期 = Format(mdateIn, "yyyy-MM-dd")
        ShowMe = mblnOK
    End If
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.dtp日期.Value = mdateIn
    
End Sub

Private Sub OKButton_Click()
    mblnOK = True
    mdateIn = Me.dtp日期.Value
    Unload Me
End Sub
