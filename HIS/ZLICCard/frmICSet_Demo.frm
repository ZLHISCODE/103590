VERSION 5.00
Begin VB.Form frmICSet_Demo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IC卡设备设置"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "frmICSet_Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraMW_RD 
      Caption         =   "IC卡参数"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2175
      Begin VB.TextBox txt_MW_Len 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "10"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txt_MW_SAddr 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "32"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "长度"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "数据起始地址"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&S)"
      Height          =   350
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1100
   End
   Begin VB.ComboBox cboCom 
      Height          =   300
      ItemData        =   "frmICSet_Demo.frx":000C
      Left            =   1185
      List            =   "frmICSet_Demo.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   1230
   End
   Begin VB.Label lblSet 
      Caption         =   "通讯端口"
      Height          =   225
      Left            =   315
      TabIndex        =   5
      Top             =   270
      Width           =   735
   End
End
Attribute VB_Name = "frmICSet_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCardType As Integer '设备编码

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveSetting "ZLSOFT", "公共全局\ICCard\" & mintCardType, "端口", cboCom.ListIndex
    If fraMW_RD.Enabled Then
        If Val(txt_MW_Len.Text) > 16 Then
            MsgBox "最大允许长度为16", vbInformation
            If txt_MW_Len.Enabled And txt_MW_Len.Visible Then txt_MW_Len.SetFocus
            Exit Sub
        End If
        
        SaveSetting "ZLSOFT", "公共全局\ICCard\" & mintCardType, "起始地址", Val(txt_MW_SAddr.Text)
        SaveSetting "ZLSOFT", "公共全局\ICCard\" & mintCardType, "长度", Val(txt_MW_Len.Text)
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call SendKeys("{Tab}")
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    cboCom.Clear
    With cboCom
        .AddItem "Com1"
        .AddItem "Com2"
        .AddItem "Com3"
        .AddItem "Com4"
        .AddItem "Com5"
        .AddItem "Com6"
        .AddItem "Com7"
        .AddItem "Com8"
        If mintCardType = 8 Then
            .AddItem "USB"
        End If
    End With
    cboCom.ListIndex = 0
    i = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mintCardType, "端口", 0))
    If i > 0 And i <= cboCom.ListCount Then cboCom.ListIndex = i

    If mintCardType = 4 Then
        fraMW_RD.Enabled = True
        txt_MW_SAddr.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mintCardType, "起始地址", 32))
        txt_MW_Len.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mintCardType, "长度", 10))
    ElseIf mintCardType = 8 Then
        fraMW_RD.Enabled = True
        txt_MW_SAddr.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mintCardType, "起始地址", 1))
        txt_MW_Len.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mintCardType, "长度", 10))
    End If
End Sub


Public Sub ShowMe(ByVal intCardType As Integer)
    mintCardType = intCardType
    Me.Show vbModal
End Sub

Private Sub txt_MW_Len_GotFocus()
    txt_MW_Len.SelStart = 0
    txt_MW_Len.SelLength = Len(txt_MW_Len)
End Sub

Private Sub txt_MW_Len_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt_MW_SAddr_GotFocus()
    txt_MW_SAddr.SelStart = 0
    txt_MW_SAddr.SelLength = Len(txt_MW_SAddr)
End Sub

Private Sub txt_MW_SAddr_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
