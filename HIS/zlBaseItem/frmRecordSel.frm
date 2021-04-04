VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form frmRecordSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6645
   Icon            =   "frmRecordSel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid dbg 
      Height          =   3090
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   5450
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1.5
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """kl ""HH.mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5310
      TabIndex        =   1
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5310
      TabIndex        =   0
      Top             =   630
      Width           =   1100
   End
End
Attribute VB_Name = "frmRecordSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnSecceed As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Me.Hide
End Sub



Private Sub cmdOK_Click()
    mblnSecceed = True
    Me.Hide
End Sub

Private Sub dbg_DblClick()
cmdOK_Click
End Sub

Private Sub Form_Activate()
    dbg.Columns(0).Width = 0
    dbg.Columns(0).Visible = False
    dbg.SetFocus
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbg.Top = 100
    dbg.Left = 100
    dbg.Height = ScaleHeight - 200
    If ScaleWidth > 3000 Then
        cmdOK.Left = ScaleWidth - cmdOK.Width - 200
        cmdCancel.Left = cmdOK.Left
        dbg.Width = cmdOK.Left - dbg.Left - 200
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub
