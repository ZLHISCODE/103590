VERSION 5.00
Begin VB.Form frmIC��֧����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ָ��IC��֧�����"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   2472
      TabIndex        =   6
      Top             =   1815
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   1302
      TabIndex        =   5
      Top             =   1815
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   -30
      TabIndex        =   4
      Top             =   1650
      Width           =   4950
   End
   Begin VB.TextBox txt֧����� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2385
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1230
      Width           =   1725
   End
   Begin VB.Label lbl�Ը���� 
      AutoSize        =   -1  'True
      Caption         =   "����������Ը�����Ϊ[0.00Ԫ]"
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Top             =   240
      Width           =   2520
   End
   Begin VB.Label lblMSG 
      AutoSize        =   -1  'True
      Caption         =   "ָ��IC��֧����"
      Height          =   180
      Left            =   765
      TabIndex        =   2
      Top             =   1305
      Width           =   1620
   End
   Begin VB.Label lblIC����� 
      AutoSize        =   -1  'True
      Caption         =   "IC�����Ϊ[0.00Ԫ]"
      Height          =   180
      Left            =   300
      TabIndex        =   1
      Top             =   950
      Width           =   1620
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      Caption         =   "���ĸ����ʻ����Ϊ[0.00Ԫ]"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   595
      Width           =   2340
   End
End
Attribute VB_Name = "frmIC��֧�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function get_IC֧�����(cur�Ը���� As Currency, cur������� As Currency, curIC����� As Currency) As Currency
    lbl�Ը����.Caption = "����������Ը�����Ϊ[" & Format(cur�Ը����, "0.00") & "Ԫ]"
    lbl�Ը����.Tag = cur�Ը����
    lbl�������.Caption = "���ĸ����ʻ����Ϊ[" & Format(cur�������, "0.00") & "Ԫ]"
    lbl�������.Tag = cur�������
    lblIC�����.Caption = "IC�����Ϊ[" & Format(curIC�����, "0.00") & "Ԫ]"
    lblIC�����.Tag = curIC�����
    
    If cur�Ը���� < cur������� + curIC����� Then
        txt֧�����.Text = Format(cur�Ը����, "0.00")
    Else
        txt֧�����.Text = Format(cur������� + curIC�����, "0.00")
    End If
    txt֧�����.Tag = cur������� + curIC�����
    
    Me.Show vbModal
    get_IC֧����� = CCur(txt֧�����.Text)
End Function

Private Sub cmdCancel_Click()
    txt֧�����.Text = -1
End Sub

Private Sub cmdOK_Click()
    txt֧�����.SetFocus
    If Not IsNumeric(txt֧�����.Text) Then
        MsgBox "������IC��֧����", vbInformation, Me.Caption
        Exit Sub
    End If
    If CCur(txt֧�����.Text) < 0 Then
        MsgBox "IC��֧������Ϊ������", vbInformation, Me.Caption
        Exit Sub
    End If
    If Len(Format(txt֧�����.Text, "0.00")) > 12 Then
        MsgBox "֧��������������ӦС��12λ��С������ӦС��2λ��", vbInformation, Me.Caption
        Exit Sub
    End If
    If CCur(txt֧�����.Text) > CCur(txt֧�����.Tag) Then
        MsgBox "֧�������������ʻ������IC�����ĺ͡�", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub txt֧�����_GotFocus()
    txt֧�����.SelStart = 0
    txt֧�����.SelLength = Len(txt֧�����.Text)
End Sub
