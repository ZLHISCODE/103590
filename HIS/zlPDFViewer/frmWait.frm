VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ӡ"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text 
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   3480
      TabIndex        =   4
      Text            =   "100%"
      Top             =   1080
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Timer timC 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   200
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "�ļ���С"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lblP 
      AutoSize        =   -1  'True
      Caption         =   "��ӡ��"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   540
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrFilePath As String
Private mlngCur As Long '����
Private mlngTim As Long '��ʱ��
Private mlng��� As Long '��ǰ��ӡ�ڼ���
Private mlng��ʱ As Long
Private mstrName As String

Public Function ShowMe(ByRef mfrmObj As Object, ByVal lng��� As Long, ByVal strFilePath As String, ByVal strName As String) As String
'���ܣ�
'������strFilePath �ļ�����·��
    mlngCur = 0
    mlngTim = 0
    mlng��� = lng���
    mstrFilePath = strFilePath
    mlng��ʱ = 250
    mstrName = strName
    Me.Show 1, mfrmObj
End Function

Private Sub Form_Load()
    Dim dblLen As Double
    Dim i As Long
 
    mlngTim = mlng��ʱ
    dblLen = FileLen(mstrFilePath) / 1024
    lblSize.Caption = "�ļ���С:" & Round(dblLen) & "KB"
    lblFileName.Caption = "������:" & mstrName
    lblP.Caption = "��ӡ��" ' "���ڴ�ӡ���档"
    timC.Interval = 1
    timC.Enabled = True '1����ִ��һ��ʱ��
End Sub

Private Sub timC_Timer()
    If mlngCur = mlngTim Then
        Progress = 100
        Progress = 0
        Unload Me
    End If
    mlngCur = mlngCur + 1
    If mlngCur < mlngTim Then
        Progress = (mlngCur / mlngTim) * 100
    End If
    If 10 = mlngCur / 1000 Then
        Unload Me
    End If
End Sub

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.Value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.Value = vNewValue
        txtPer.Text = CInt(psb.Value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

