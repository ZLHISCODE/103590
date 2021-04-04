VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印设置"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3285
      TabIndex        =   4
      Top             =   1200
      Width           =   1100
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3390
   End
   Begin VB.TextBox txtCopy 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   960
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   660
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "打 印 机："
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "打印份数："
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public blnOk As Boolean


'显示打印机设置
Public Sub ShowPrinterSet(owner As Object)
    Me.Show 1, owner
End Sub


Private Sub cmdCancel_Click()
    blnOk = False
    
    Me.Hide
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim i As Long
    
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = cboPrinter.Text Then
            Set Printer = Printers(i)
            Exit For
        End If
    Next
    
    Printer.Copies = Val(txtCopy.Text)
    
    blnOk = True
    
    Me.Hide
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim i As Long
    
    For i = 0 To Printers.Count - 1
        cboPrinter.AddItem Printers(i).DeviceName
        
        If Printers(i).DeviceName = Printer.DeviceName Then
            cboPrinter.ListIndex = cboPrinter.ListCount - 1
            txtCopy.Text = Printer.Copies
        End If
    Next
    
    blnOk = False
    
    Err.Clear
End Sub
