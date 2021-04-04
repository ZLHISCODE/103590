VERSION 5.00
Begin VB.Form frm病种查找 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frm病种查找.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2220
      TabIndex        =   4
      Top             =   900
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   960
      TabIndex        =   3
      Top             =   900
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   2
      Top             =   690
      Width           =   5085
   End
   Begin VB.TextBox txt内容 
      Height          =   300
      Left            =   870
      TabIndex        =   1
      Top             =   210
      Width           =   2655
   End
   Begin VB.Label lbl内容 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "内容"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   270
      Width           =   360
   End
End
Attribute VB_Name = "frm病种查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lng险类 As Long
Private rsFind As New ADODB.Recordset

Private Sub cmdFind_Click()
    Dim strInput As String
    strInput = Trim(UCase(txt内容.Text))
    
    If Val(cmdFind.Tag) = 0 Then
        gstrSQL = " Select ID From 保险病种 " & _
                  " Where (编码 Like '%" & strInput & "%' Or 名称 Like '%" & strInput & "%') And 险类=" & lng险类
        Call OpenRecordset(rsFind, "获取查找记录集")
        If rsFind.RecordCount = 0 Then
            MsgBox "没有找到任何记录，请重输！", vbInformation, gstrSysName
            txt内容.SetFocus
            Exit Sub
        Else
            cmdFind.Tag = 1
            cmdFind.Caption = "下一条(&N)"
        End If
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            If MsgBox("你需要从头开始查找吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txt内容.SetFocus
                cmdFind.Tag = 0
                cmdFind.Caption = "查找(&F)"
                Exit Sub
            Else
                rsFind.MoveFirst
            End If
        End If
    End If
    
    '根据记录集定位主界面
    With frm保险病种
        Dim lngRow As Long, lngItems As Long
        lngItems = .lvwItem.ListItems.Count
        For lngRow = 1 To lngItems
            If Val(Mid(.lvwItem.ListItems(lngRow).Key, 2)) = rsFind!ID Then
                .lvwItem.ListItems(lngRow).Selected = True
                .lvwItem.SelectedItem.Selected = True
                .lvwItem.SelectedItem.EnsureVisible
                Exit For
            End If
        Next
    End With
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt内容_GotFocus()
    Call zlControl.TxtSelAll(txt内容)
End Sub
