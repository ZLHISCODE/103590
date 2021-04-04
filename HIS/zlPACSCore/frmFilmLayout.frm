VERSION 5.00
Begin VB.Form frmFilmLayout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "图像分格"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "frmFilmLayout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2820
      TabIndex        =   14
      Top             =   4170
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   480
      TabIndex        =   13
      Top             =   4170
      Width           =   1100
   End
   Begin VB.Frame famImage 
      Height          =   3930
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.CheckBox chkClearAfterPrint 
         Caption         =   "打印后清空"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CheckBox chkCheckPrinter 
         Caption         =   "检测打印机状态"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   7
         Left            =   2100
         Picture         =   "frmFilmLayout.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1170
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   6
         Left            =   1200
         Picture         =   "frmFilmLayout.frx":1ABE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   5
         Left            =   300
         Picture         =   "frmFilmLayout.frx":3570
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   4
         Left            =   3000
         Picture         =   "frmFilmLayout.frx":5022
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   3
         Left            =   2100
         Picture         =   "frmFilmLayout.frx":6AD4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   2
         Left            =   1200
         Picture         =   "frmFilmLayout.frx":8586
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   1
         Left            =   300
         Picture         =   "frmFilmLayout.frx":A038
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   276
         Width           =   900
      End
      Begin VB.Frame Frame4 
         Caption         =   "自定义"
         Height          =   1185
         Left            =   330
         TabIndex        =   1
         Top             =   2160
         Width           =   3615
         Begin VB.CommandButton CmdApply 
            Caption         =   "应用(&A)"
            Height          =   350
            Left            =   1320
            TabIndex        =   15
            Top             =   690
            Width           =   1100
         End
         Begin VB.ListBox lstCol 
            Height          =   240
            ItemData        =   "frmFilmLayout.frx":BAEA
            Left            =   2400
            List            =   "frmFilmLayout.frx":BAFD
            TabIndex        =   12
            Top             =   270
            Width           =   885
         End
         Begin VB.ListBox lstRow 
            Height          =   240
            ItemData        =   "frmFilmLayout.frx":BB10
            Left            =   600
            List            =   "frmFilmLayout.frx":BB23
            TabIndex        =   11
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "行:"
            Height          =   180
            Left            =   240
            TabIndex        =   3
            Top             =   300
            Width           =   270
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "列:"
            Height          =   180
            Left            =   2040
            TabIndex        =   2
            Top             =   300
            Width           =   270
         End
      End
   End
End
Attribute VB_Name = "frmFilmLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrFilmLayout As String

Private Sub cmdApply_Click()
    mstrFilmLayout = Me.lstCol.list(Me.lstCol.TopIndex) & "," & Me.lstRow.list(Me.lstRow.TopIndex)
    Unload Me
End Sub

Private Sub CmdCancel_Click()
    mstrFilmLayout = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
Public Function ShowMe(objfrm As Object) As String
    Me.Show vbModal, objfrm
    ShowMe = mstrFilmLayout
End Function

Private Sub CmdSerial_Click(Index As Integer)
    Select Case Index
        Case 1
            mstrFilmLayout = "1,1"
        Case 2
            mstrFilmLayout = "2,1"
        Case 3
            mstrFilmLayout = "1,2"
        Case 4
            mstrFilmLayout = "2,2"
        Case 5
            mstrFilmLayout = "4,2"
        Case 6
            mstrFilmLayout = "2,4"
        Case 7
            mstrFilmLayout = "4,4"
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intTemp As Integer
    
    intTemp = Val(GetSetting("ZLSOFT", "公共模块\zlPacsCore", "检查打印机状态", "0"))
    If intTemp = 1 Then
        chkCheckPrinter.Value = 1
    Else
        chkCheckPrinter.Value = 0
    End If
    
    '打印后清空，默认值是1
    intTemp = Val(GetSetting("ZLSOFT", "公共模块\zlPacsCore", "打印后清空", "1"))
    If intTemp = 0 Then
        chkClearAfterPrint.Value = 0
    Else
        chkClearAfterPrint.Value = 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "公共模块\zlPacsCore", "检查打印机状态", chkCheckPrinter.Value
    SaveSetting "ZLSOFT", "公共模块\zlPacsCore", "打印后清空", chkClearAfterPrint.Value
End Sub
