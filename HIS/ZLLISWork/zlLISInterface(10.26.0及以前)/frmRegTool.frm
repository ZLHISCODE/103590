VERSION 5.00
Begin VB.Form frmRegTool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������LIS�ӿ���Ȩ�����ɹ���"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   435
      Left            =   960
      TabIndex        =   3
      Top             =   1215
      Width           =   4320
   End
   Begin VB.CommandButton cmdGetReg 
      Caption         =   "������Ȩ�ļ�"
      Height          =   480
      Left            =   3495
      TabIndex        =   2
      Top             =   1905
      Width           =   1620
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "3000-01-01"
      Top             =   405
      Width           =   4290
   End
   Begin VB.TextBox txtUnti 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Text            =   "����ҽԺ��Ϣ����ϵͳ"
      Top             =   780
      Width           =   4290
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ��"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   6
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ�û�"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ֹ����"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   435
      Width           =   720
   End
End
Attribute VB_Name = "frmRegTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetReg_Click()
    If Trim(txtUnti.Text) = "" Then
        MsgBox "��Ȩ�û�����Ϊ��!"
        Exit Sub
    End If
    If Trim(txtDate.Text) = "" Then
        MsgBox "��Ȩ��ֹ���ڲ���Ϊ��!"
        Exit Sub
    End If
    If Not IsDate(txtDate.Text) Or Not (txtDate.Text Like "####-##-##") Then
        MsgBox "��Ȩ��ֹ���ڸ�ʽ����ȷ���밴yyyy-MM-dd��ʽ����!"
        Exit Sub
    End If
    Call GetRegInfo
End Sub
Private Function GetRegInfo() As Boolean

    Dim strLine As String, str���� As String, date���� As Date
    Dim strUnti As String, strCode As String, strKey As String
    Dim lngFile  As Long
    
    strKey = "�¶�"
    strUnti = txtUnti.Text
    str���� = txtDate.Text
    strCode = sha1(EncodeBase64String(strUnti & "|" & str���� & "|" & strKey))
    txtCode = strCode
    lngFile = FreeFile
    Open App.Path & "\RegFile.ini" For Output Access Write As lngFile
    
    Print #lngFile, "��Ȩ��ֹ����=" & str���� & vbNewLine & "��Ȩ��=" & strCode
    Close lngFile
    MsgBox "��������Ȩ�ļ���:" & App.Path & "\RegFile.ini"
End Function

