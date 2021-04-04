VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间输入"
   ClientHeight    =   2085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4875
   Icon            =   "frmInputDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1800
      Left            =   75
      TabIndex        =   3
      Top             =   60
      Width           =   3405
      Begin VB.CheckBox chk 
         Caption         =   "分娩"
         Height          =   300
         Index           =   1
         Left            =   1950
         TabIndex        =   6
         Top             =   1215
         Width           =   795
      End
      Begin VB.CheckBox chk 
         Caption         =   "手术"
         Height          =   300
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   1215
         Value           =   1  'Checked
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   735
         TabIndex        =   0
         Top             =   765
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   89456643
         CurrentDate     =   38952
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   180
         Picture         =   "frmInputDate.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "请在下面输入有效的手术时间"
         Height          =   180
         Left            =   690
         TabIndex        =   4
         Top             =   285
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3600
      TabIndex        =   2
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      Top             =   165
      Width           =   1100
   End
End
Attribute VB_Name = "frmInputDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mstrTime As String
Private mstrCaption As String

Public Function ShowMe(ByRef strTime As String, ByVal strMin As String, ByVal strMax As String, ByRef strCaption As String) As Boolean
    
    mblnOK = False
    mstrTime = strTime
    mstrCaption = strCaption
    
    dtp(0).MinDate = Format(strMin, dtp(0).CustomFormat)
    dtp(0).MaxDate = Format(strMax, dtp(0).CustomFormat)
    dtp(0).Value = Format(mstrTime, dtp(0).CustomFormat)
    
    Select Case mstrCaption
    Case "手术"
        chk(0).Value = 1
        chk(1).Value = 0
    Case "分娩"
        chk(0).Value = 0
        chk(1).Value = 1
    Case "手术分娩"
        chk(0).Value = 1
        chk(1).Value = 1
    End Select
    
    Me.Show 1
    strCaption = mstrCaption
    strTime = mstrTime
    ShowMe = mblnOK
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If chk(0).Value = 0 And chk(1).Value = 0 Then
        ShowSimpleMsg "至少要选择一项内容，手术或分娩！"
        Exit Sub
    End If
    
    mstrTime = Format(dtp(0).Value, dtp(0).CustomFormat)
    
    If mstrTime < Format(dtp(0).MinDate, dtp(0).CustomFormat) Or mstrTime > Format(dtp(0).MaxDate, dtp(0).CustomFormat) Then
        ShowSimpleMsg "输入的时间不在范围内！"
        Exit Sub
    End If
    
    mstrCaption = ""
    If chk(0).Value = 1 Then mstrCaption = "手术"
    If chk(1).Value = 1 Then mstrCaption = mstrCaption & "分娩"
    
    mblnOK = True
    Unload Me
End Sub

