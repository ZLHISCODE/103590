VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetCustomTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置自定义时间范围"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   Icon            =   "frmSetCustomTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   240
      Picture         =   "frmSetCustomTime.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   120
      Width           =   500
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   800
      Width           =   4935
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1880
      Width           =   5055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   3480
      TabIndex        =   2
      Top             =   2040
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy""年""MM""月""dd""日"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   233570307
      CurrentDate     =   36257
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy""年""MM""月""dd""日"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   233570307
      CurrentDate     =   36257
   End
   Begin VB.Label lbl 
      Caption         =   "结束时间"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   1470
      Width           =   840
   End
   Begin VB.Label lbl 
      Caption         =   "开始时间"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   990
      Width           =   840
   End
   Begin VB.Label lbl 
      Caption         =   "请设置自定义时间范围的开始时间和结束时间"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   3600
   End
End
Attribute VB_Name = "frmSetCustomTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Dim mstrTimeBegin As String
Dim mstrTimeEnd As String

Public Function ShowMe(ByVal frmParent As Object, ByRef strTimeBegin As String, ByRef strTimeEnd As String) As String
    mstrTimeBegin = strTimeBegin
    mstrTimeEnd = strTimeEnd
    dtpEnd.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpBegin.MaxDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59")
    dtpBegin.Value = Format(CDate(strTimeBegin), "yyyy-mm-dd 00:00:00")
    dtpEnd.Value = Format(CDate(strTimeEnd), "yyyy-mm-dd 23:59:59")
    Me.Show 1, frmParent
    strTimeBegin = mstrTimeBegin
    strTimeEnd = mstrTimeEnd
End Function

Private Sub cmdOK_Click()
    mstrTimeBegin = Format(dtpBegin.Value, "yyyy-mm-dd HH:mm:ss")
    mstrTimeEnd = Format(dtpEnd.Value, "yyyy-mm-dd HH:mm:ss")
    Unload Me
End Sub

Private Sub dtpBegin_Validate(Cancel As Boolean)
    dtpEnd.MinDate = dtpBegin.Value
End Sub

Private Sub dtpEnd_Validate(Cancel As Boolean)
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
