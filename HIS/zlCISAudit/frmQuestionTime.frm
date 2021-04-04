VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQuestionTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间"
   ClientHeight    =   1800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4035
   Icon            =   "frmQuestionTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1500
      TabIndex        =   4
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2595
      TabIndex        =   5
      Top             =   1260
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   6
      Top             =   1050
      Width           =   4440
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   2010
      TabIndex        =   3
      Top             =   570
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   72744963
      CurrentDate     =   39158
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   2010
      TabIndex        =   1
      Top             =   150
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   72744963
      CurrentDate     =   39158
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      Height          =   180
      Left            =   1200
      TabIndex        =   0
      Top             =   210
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束时间"
      Height          =   180
      Left            =   1200
      TabIndex        =   2
      Top             =   630
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   255
      Picture         =   "frmQuestionTime.frx":000C
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmQuestionTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK          As Boolean
Private mdBegin         As Date
Private mdEnd           As Date

Public Function ShowMe(frmParent As Object, dBegin As Date, dEnd As Date) As Boolean
    mdBegin = dBegin
    mdEnd = dEnd
    Me.Show 1, frmParent
    
    If mblnOK Then
        dBegin = mdBegin
        dEnd = mdEnd
    End If
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
       MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
       dtpBegin.SetFocus: Exit Sub
    End If
    
    mdBegin = Format(dtpBegin.Value, "yyyy-MM-dd 00:00:00")
    mdEnd = Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
    
    dtpEnd.MaxDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.MaxDate = dtpEnd.MaxDate
    
    If mdBegin = CDate(0) Or mdEnd = CDate(0) Then
        '缺省为当天
        dtpBegin.Value = Format(dtpEnd.MaxDate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = dtpEnd.MaxDate
    Else
        dtpBegin.Value = mdBegin
        dtpEnd.Value = mdEnd
    End If
End Sub

