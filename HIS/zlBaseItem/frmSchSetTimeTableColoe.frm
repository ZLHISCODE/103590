VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSchSetTimeTableColoe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ԤԼ--ʱ�����ɫ����"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmSchSetTimeTableColoe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   720
      TabIndex        =   5
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "��"
      Height          =   255
      Index           =   4
      Left            =   3975
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "��"
      Height          =   255
      Index           =   3
      Left            =   3975
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "��"
      Height          =   255
      Index           =   2
      Left            =   3975
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "��"
      Height          =   255
      Index           =   0
      Left            =   3975
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "��"
      Height          =   255
      Index           =   1
      Left            =   3975
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lab 
      Caption         =   "ԤԼ��ǩ��ɫ���ѹ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   2880
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "ԤԼ��ǩ��ɫ�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2880
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "ԤԼ��ǩ��ɫ����ԤԼ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2880
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "ʱ�����ɫ������ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2880
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "ʱ�����ɫ����Ϣʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   2295
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2880
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmSchSetTimeTableColoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngColorTabRest As Long    'ԤԼʱ�����Ϣʱ����ɫ
Private mlngColorTabWork As Long    'ԤԼʱ�������ʱ����ɫ
Private mlngColorLblWaiting As Long 'ԤԼ��ǩ��ԤԼ�Ⱥ���ɫ
Private mlngColorLblDone As Long    'ԤԼ��ǩ�������ɫ
Private mlngColorLblPassed As Long  'ԤԼ��ǩ��������ɫ

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click(Index As Integer)
    dlgColor.Color = shpColor(Index).FillColor
    dlgColor.ShowColor
    shpColor(Index).FillColor = dlgColor.Color
End Sub

Private Sub cmdOK_Click()
    
    Call zlDatabase.SetPara("���ԤԼʱ�����ʱ����ɫ", shpColor(0).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("���ԤԼʱ�����Ϣʱ����ɫ", shpColor(1).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("���ԤԼ��ǩ��ԤԼ��ɫ", shpColor(2).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("���ԤԼ��ǩ�������ɫ", shpColor(3).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("���ԤԼ��ǩ�ѹ�����ɫ", shpColor(4).FillColor, glngSys, 1292)
    Unload Me
End Sub

Private Sub Form_Load()
        
    Call LoadColors
    
    shpColor(0).FillColor = mlngColorTabWork
    shpColor(1).FillColor = mlngColorTabRest
    shpColor(2).FillColor = mlngColorLblWaiting
    shpColor(3).FillColor = mlngColorLblDone
    shpColor(4).FillColor = mlngColorLblPassed
    
End Sub

Private Sub LoadColors()
'------------------------------------------------
'���ܣ�װ��ʱ���ı���ʽ�ͻ�������
'������
'���أ���
'------------------------------------------------
    
    On Error GoTo err
    
    '�����ݿ��ж�ȡ���ù�����ɫ
    mlngColorTabWork = zlDatabase.GetPara("���ԤԼʱ�����ʱ����ɫ", glngSys, 1292, "8421376")
    mlngColorTabRest = zlDatabase.GetPara("���ԤԼʱ�����Ϣʱ����ɫ", glngSys, 1292, "16777215")
    mlngColorLblWaiting = zlDatabase.GetPara("���ԤԼ��ǩ��ԤԼ��ɫ", glngSys, 1292, "0")
    mlngColorLblDone = zlDatabase.GetPara("���ԤԼ��ǩ�������ɫ", glngSys, 1292, "12632256")
    mlngColorLblPassed = zlDatabase.GetPara("���ԤԼ��ǩ�ѹ�����ɫ", glngSys, 1292, "255")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

