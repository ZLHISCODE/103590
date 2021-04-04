VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelectTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ʱ��ѡ��"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "frmSelectTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -45
      TabIndex        =   6
      Top             =   1155
      Width           =   4440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2550
      TabIndex        =   3
      Top             =   1365
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1455
      TabIndex        =   2
      Top             =   1365
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1965
      TabIndex        =   1
      Top             =   675
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   166920195
      CurrentDate     =   39158
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1965
      TabIndex        =   0
      Top             =   255
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   166920195
      CurrentDate     =   39158
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   210
      Picture         =   "frmSelectTime.frx":058A
      Top             =   165
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   1155
      TabIndex        =   5
      Top             =   735
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼʱ��"
      Height          =   180
      Left            =   1155
      TabIndex        =   4
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmSelectTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mdBegin As Date
Private mdEnd As Date
Private mlngX, mlngY As Long '��ʾ������λ��

Public Function ShowMe(frmParent As Object, dBegin As Date, dEnd As Date, ByVal objControl As Variant) As Boolean
'���ܣ���ʾһ��ʱ��ѡ���(����)������ѡ��ʱ�䷶Χ
'      ���� objControl �ǿؼ����֣���������������ʾλ��
    Dim vPoint As POINTAPI
    mdBegin = dBegin
    mdEnd = dEnd
    
    vPoint = gobjComlib.zlControl.GetCoordPos(objControl.hWnd, objControl.Width, objControl.Height)
    mlngX = vPoint.X: mlngY = vPoint.Y
    
    Me.Show 1, frmParent
    
    If mblnOk Then
        dBegin = mdBegin
        dEnd = mdEnd
    End If
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
       MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
       dtpBegin.SetFocus: Exit Sub
    End If
    
    mdBegin = Format(dtpBegin.Value, "yyyy-MM-dd 00:00:00")
    mdEnd = Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")
    
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call gobjComlib.ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOk = False
    dtpEnd.MaxDate = Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.MaxDate = dtpEnd.MaxDate
    
    If mlngX <= 0 Or mlngY <= 0 Then '���곬����Ļ�⣬��ʾ������
        SetWindowPos frmSelectTime.hWnd, HWND_TOPMOST, Screen.Width / 2, Screen.Height / 2, 0, 0, &H10 Or &H1
    Else
        SetWindowPos frmSelectTime.hWnd, HWND_TOPMOST, mlngX / Screen.TwipsPerPixelX, mlngY / Screen.TwipsPerPixelY, 0, 0, &H10 Or &H1
    End If
    
    If mdBegin = CDate(0) Or mdEnd = CDate(0) Then
        'ȱʡΪ����
        dtpBegin.Value = Format(dtpEnd.MaxDate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = dtpEnd.MaxDate
    Else
        dtpBegin.Value = mdBegin
        dtpEnd.Value = mdEnd
    End If
End Sub
