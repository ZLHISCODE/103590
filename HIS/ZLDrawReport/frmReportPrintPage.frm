VERSION 5.00
Begin VB.Form frmReportPrintPage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印选项"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "frmReportPrintPage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Caption         =   "页码计算方法"
      Height          =   735
      Index           =   1
      Left            =   105
      TabIndex        =   15
      Top             =   2370
      Width           =   5775
      Begin VB.OptionButton optStyle 
         Appearance      =   0  'Flat
         Caption         =   "按标记的页码计算（即排除封面和目录）"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2130
         TabIndex        =   11
         Top             =   315
         Value           =   -1  'True
         Width           =   3540
      End
      Begin VB.OptionButton optStyle 
         Appearance      =   0  'Flat
         Caption         =   "按实际页数计算"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   315
         Width           =   1890
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Enabled         =   0   'False
         Height          =   180
         Index           =   3
         Left            =   2970
         TabIndex        =   16
         Top             =   1125
         Width           =   90
      End
   End
   Begin VB.Frame fra 
      Caption         =   "打印范围"
      Height          =   2250
      Index           =   0
      Left            =   105
      TabIndex        =   14
      Top             =   60
      Width           =   5775
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   1545
         TabIndex        =   2
         Top             =   690
         Width           =   2805
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "指定页码"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   705
         Width           =   1125
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "页码范围"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1095
         Width           =   1140
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   2115
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   3540
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1080
         Width           =   810
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "偶数页码"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1875
         Width           =   1470
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "奇数页码"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1500
         Width           =   1200
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "全部(&A)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   315
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "从(&F):"
         Enabled         =   0   'False
         Height          =   180
         Index           =   0
         Left            =   1545
         TabIndex        =   4
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "到(&T):"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   2970
         TabIndex        =   6
         Top             =   1125
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3570
      TabIndex        =   12
      Top             =   3165
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4740
      TabIndex        =   13
      Top             =   3165
      Width           =   1100
   End
End
Attribute VB_Name = "frmReportPrintPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPage As String
Private mblnOK As Boolean
Private mbytPageStype As Byte

Public Function ShowDialog(ByVal objParent As Object, ByRef bytPageStype As Byte, ByRef strPage As String) As Boolean
    
    mblnOK = False
    
    Me.Show 1, objParent
    
    If mblnOK Then
        strPage = mstrPage
        bytPageStype = mbytPageStype
    End If
    
    ShowDialog = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Call Unload(Me)
End Sub

Private Sub cmdOK_Click()
    Dim lngPage As Long
    
    mstrPage = ""
    
    If opt(0).value = True Then mstrPage = ""
    If opt(1).value = True Then mstrPage = txt(2).Text
    If opt(2).value = True Then
        
        For lngPage = Val(txt(0).Text) To Val(txt(1).Text)
            mstrPage = mstrPage & "," & lngPage
        Next
        
        If mstrPage <> "" Then mstrPage = Mid(mstrPage, 2)
    End If
    
    If opt(3).value = True Then mstrPage = "-1"
    If opt(4).value = True Then mstrPage = "-2"
    
    mbytPageStype = IIf(optStyle(0).value = True, 2, 1)
    
    mblnOK = True
    Call Unload(Me)
End Sub

Private Sub opt_Click(Index As Integer)
    txt(0).Enabled = (opt(2).value = True)
    txt(1).Enabled = (opt(2).value = True)
    txt(2).Enabled = (opt(1).value = True)
    lbl(0).Enabled = (opt(2).value = True)
    lbl(1).Enabled = (opt(2).value = True)
    
    txt(0).BackColor = IIf(opt(2).value = True, &H80000005, &H8000000F)
    txt(1).BackColor = IIf(opt(2).value = True, &H80000005, &H8000000F)
    txt(2).BackColor = IIf(opt(1).value = True, &H80000005, &H8000000F)
    
    Select Case Index
    Case 1
        DoEvents
        txt(2).SetFocus
    Case 2
        DoEvents
        txt(0).SetFocus
    End Select
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub optStyle_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        OS.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then
            If FilterKeyAscii(KeyAscii, 99, "0123456789,") = 0 Then KeyAscii = 0
        End If
        
        If Index = 0 Or Index = 1 Then
            If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function
