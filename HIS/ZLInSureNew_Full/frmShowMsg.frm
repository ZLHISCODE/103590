VERSION 5.00
Begin VB.Form frmShowMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ҽ�ƻ����˶���Ϣ"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
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
   Begin VB.Label lblҽ�ƻ��� 
      AutoSize        =   -1  'True
      Caption         =   "��¼����:"
      Height          =   180
      Index           =   0
      Left            =   4305
      TabIndex        =   3
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��¼����:"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�ƻ���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
    '��ʾ��Ϣ,������|ҽ�ƻ���||����|ҽ�ƻ���
    Dim strArr
    Dim strArr1
    Dim i As Long
    strArr = Split(strInfor, "||")
    For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), "|")
        If i > 0 Then
            Load lbl����(i)
            Load lblҽ�ƻ���(i)
        End If
        With lbl����(i)
            .Visible = True
            .Left = lbl����(0).Left
            If i > 0 Then
                .Top = lbl����(i - 1).Top + lbl����(i - 1).Height + 100
            End If
            .Caption = strArr1(0)
        End With
        
        With lblҽ�ƻ���(i)
            .Visible = True
            .Left = lblҽ�ƻ���(0).Left
            If i > 0 Then
                .Top = lblҽ�ƻ���(i - 1).Top + lblҽ�ƻ���(i - 1).Height + 100
            End If
            .Caption = strArr1(1)
        End With
        
    Next
    With fra(1)
        .Top = lbl����(UBound(strArr)).Top + 400
        Me.CMD����.Top = .Top + .Height + 100
        Me.Height = Me.CMD����.Height + Me.CMD����.Top + 500
    End With
    Me.Show 1
End Sub
     
Private Sub CMD����_Click()

    Unload Me
End Sub

