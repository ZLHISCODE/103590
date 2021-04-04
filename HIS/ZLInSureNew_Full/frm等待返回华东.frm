VERSION 5.00
Begin VB.Form frm等待返回华东 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "等待返回......"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer TimeRead 
      Interval        =   1000
      Left            =   270
      Top             =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   4080
      TabIndex        =   3
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2940
      TabIndex        =   2
      Top             =   1320
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   23
      Picture         =   "frm等待返回华东.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   5325
   End
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   15
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "等待医保返回结算数据......"
      Height          =   180
      Left            =   1515
      TabIndex        =   1
      Top             =   450
      Width           =   2340
   End
End
Attribute VB_Name = "frm等待返回华东"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String, isOK As Boolean

Public Function waitReturn(strFile As String) As Boolean
    mstrFile = strFile
    Me.Show vbModal
    waitReturn = isOK
End Function

Private Sub cmdCancel_Click()
    isOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    isOK = True
    Me.Hide
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub

Private Sub TimeRead_Timer()
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    Set rsTemp = gcn华东.Execute("Select * From " & mstrFile)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    If rsTemp.EOF Then Exit Sub
    
    On Error GoTo errHandle
    Label1.Caption = "已找到结算数据"
    cmdOK.Enabled = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
