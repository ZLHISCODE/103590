VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMedicalStationSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   3885
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5520
   Icon            =   "frmMedicalStationSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "正在体检"
      Height          =   1305
      Left            =   105
      TabIndex        =   22
      Top             =   1950
      Width           =   5355
      Begin VB.CheckBox chk 
         Caption         =   "按报到时间查(&5)"
         Height          =   240
         Index           =   2
         Left            =   900
         TabIndex        =   16
         Top             =   990
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   900
         TabIndex        =   9
         Top             =   270
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   3
         Left            =   3180
         TabIndex        =   11
         Top             =   270
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   6
         Left            =   900
         TabIndex        =   13
         Top             =   630
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   7
         Left            =   3180
         TabIndex        =   15
         Top             =   630
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "个  人(&3)"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   4
         Left            =   2985
         TabIndex        =   10
         Top             =   315
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   6
         Left            =   2985
         TabIndex        =   14
         Top             =   675
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "团  体(&4)"
         Height          =   180
         Index           =   7
         Left            =   60
         TabIndex        =   12
         Top             =   690
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3330
      Width           =   1100
   End
   Begin VB.Frame Frame5 
      Caption         =   "时间范围"
      Height          =   1665
      Left            =   105
      TabIndex        =   20
      Top             =   210
      Width           =   5340
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   915
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   3180
         TabIndex        =   3
         Top             =   915
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   4
         Left            =   900
         TabIndex        =   5
         Top             =   1275
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   5
         Left            =   3180
         TabIndex        =   7
         Top             =   1275
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   98435075
         CurrentDate     =   39000
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   5
         Left            =   2985
         TabIndex        =   6
         Top             =   1335
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   0
         Left            =   2985
         TabIndex        =   2
         Top             =   990
         Width           =   180
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmMedicalStationSearch.frx":000C
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label9 
         Caption         =   "在体检工作站中的待体检、正体检以及已完成体检的时间范围分别按如下设置进行搜索。"
         Height          =   405
         Left            =   795
         TabIndex        =   21
         Top             =   375
         Width           =   4065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "待体检(&1)"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   0
         Top             =   975
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "已完成(&2)"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   4
         Top             =   1335
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4245
      TabIndex        =   18
      Top             =   3330
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3105
      TabIndex        =   17
      Top             =   3330
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalStationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrCondition As String

Public Function ShowFilter(ByVal frmMain As Object, ByRef strCondition As String) As Boolean
    
    mblnOK = False
    mstrCondition = strCondition
    
    Set mfrmMain = frmMain
    '初始化

    dtp(0).Value = Format(Split(Split(mstrCondition, "'")(0), "|")(0), dtp(0).CustomFormat)
    dtp(1).Value = Format(Split(Split(mstrCondition, "'")(0), "|")(1), dtp(1).CustomFormat)
    dtp(2).Value = Format(Split(Split(mstrCondition, "'")(1), "|")(0), dtp(2).CustomFormat)
    dtp(3).Value = Format(Split(Split(mstrCondition, "'")(1), "|")(1), dtp(3).CustomFormat)
    dtp(4).Value = Format(Split(Split(mstrCondition, "'")(2), "|")(0), dtp(4).CustomFormat)
    dtp(5).Value = Format(Split(Split(mstrCondition, "'")(2), "|")(1), dtp(5).CustomFormat)
        
    dtp(6).Value = Format(Split(Split(mstrCondition, "'")(3), "|")(0), dtp(6).CustomFormat)
    dtp(7).Value = Format(Split(Split(mstrCondition, "'")(3), "|")(1), dtp(7).CustomFormat)
    
    chk(2).Value = Val(Split(mstrCondition, "'")(4))
    
    Me.Show 1, frmMain
    strCondition = mstrCondition
    ShowFilter = mblnOK
    
End Function

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    
    mstrCondition = Format(dtp(0).Value, dtp(0).CustomFormat)
    mstrCondition = mstrCondition & "|" & Format(dtp(1).Value, dtp(1).CustomFormat)
    
    mstrCondition = mstrCondition & "'" & Format(dtp(2).Value, dtp(2).CustomFormat)
    mstrCondition = mstrCondition & "|" & Format(dtp(3).Value, dtp(3).CustomFormat)
    
    mstrCondition = mstrCondition & "'" & Format(dtp(4).Value, dtp(4).CustomFormat)
    mstrCondition = mstrCondition & "|" & Format(dtp(5).Value, dtp(5).CustomFormat)
    
    mstrCondition = mstrCondition & "'" & Format(dtp(6).Value, dtp(6).CustomFormat)
    mstrCondition = mstrCondition & "|" & Format(dtp(7).Value, dtp(7).CustomFormat)
    
    mstrCondition = mstrCondition & "'" & chk(2).Value
    
    mblnOK = True

    
    Unload Me
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
