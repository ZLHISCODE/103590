VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTimeSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日期范围设置"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确定"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPEnd 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112787457
      CurrentDate     =   42898
   End
   Begin MSComCtl2.DTPicker DTPStart 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112787457
      CurrentDate     =   42898
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "结束时间："
      Height          =   180
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "开始时间："
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frmTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDTStart As Date
Private mDTEnd As Date
Private mdateRange As Integer


Public Function GetTimeSet(ByRef dtStart As Date, ByRef dtEnd As Date)
    dtStart = mDTStart
    dtEnd = mDTEnd
End Function

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSure_Click()
    Dim dtTmp As Date
    
    If DTPStart.value > DTPEnd.value Then
        MsgBox "开始时间必须早于结束时间，请重新设置", , "日期范围设置"
        Exit Sub
    End If
    If mdateRange > 0 Then
        dtTmp = DateAdd("yyyy", -mdateRange, DTPEnd.value)

        If dtTmp > DTPStart.value Then
            MsgBox "时间超过配置的时间范围，请重新设置", , "日期范围设置"
            Exit Sub
        End If
        
        
    End If
    
    mDTStart = DTPStart.value
    mDTEnd = DTPEnd.value
    Me.Hide
End Sub

Private Sub Form_Activate()
    DTPStart.value = mDTStart
    DTPEnd.value = mDTEnd
End Sub


Public Function zlShowMe(ByVal frmOwner As Object, ByVal dtStart As Date, ByVal dtEnd As Date, ByVal dateRange As Integer) As Boolean
    mDTStart = dtStart
    mDTEnd = dtEnd
    mdateRange = dateRange
    
    Me.Move frmOwner.Left + (frmOwner.Width - Me.Width) / 2, frmOwner.Top + (frmOwner.Height - Me.Height) / 2
    Call Me.Show(1, frmOwner)
End Function
