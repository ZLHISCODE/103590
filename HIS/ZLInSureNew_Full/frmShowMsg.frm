VERSION 5.00
Begin VB.Form frmShowMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中心与医疗机构核对信息"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmShowMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4860
      TabIndex        =   6
      Top             =   2355
      Width           =   1230
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   1
      Left            =   -15
      TabIndex        =   5
      Top             =   2055
      Width           =   7260
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   0
      Left            =   -60
      TabIndex        =   4
      Top             =   735
      Width           =   7230
   End
   Begin VB.Label lbl医疗机构 
      AutoSize        =   -1  'True
      Caption         =   "记录总数:"
      Height          =   180
      Index           =   0
      Left            =   4305
      TabIndex        =   3
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lbl中心 
      AutoSize        =   -1  'True
      Caption         =   "记录总数:"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医疗机构"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4215
      TabIndex        =   1
      Top             =   390
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "中心"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   690
      TabIndex        =   0
      Top             =   390
      Width           =   600
   End
End
Attribute VB_Name = "frmShowMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowInFor(ByVal strInfor As String)
    '显示信息,以中心|医疗机构||中心|医疗机构
    Dim strArr
    Dim strArr1
    Dim i As Long
    strArr = Split(strInfor, "||")
    For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), "|")
        If i > 0 Then
            Load lbl中心(i)
            Load lbl医疗机构(i)
        End If
        With lbl中心(i)
            .Visible = True
            .Left = lbl中心(0).Left
            If i > 0 Then
                .Top = lbl中心(i - 1).Top + lbl中心(i - 1).Height + 100
            End If
            .Caption = strArr1(0)
        End With
        
        With lbl医疗机构(i)
            .Visible = True
            .Left = lbl医疗机构(0).Left
            If i > 0 Then
                .Top = lbl医疗机构(i - 1).Top + lbl医疗机构(i - 1).Height + 100
            End If
            .Caption = strArr1(1)
        End With
        
    Next
    With fra(1)
        .Top = lbl中心(UBound(strArr)).Top + 400
        Me.CMD放弃.Top = .Top + .Height + 100
        Me.Height = Me.CMD放弃.Height + Me.CMD放弃.Top + 500
    End With
    Me.Show 1
End Sub
     
Private Sub CMD放弃_Click()

    Unload Me
End Sub

