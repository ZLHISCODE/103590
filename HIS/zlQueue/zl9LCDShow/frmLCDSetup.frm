VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLCDSetup 
   Caption         =   "��������"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   Icon            =   "frmLCDSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   5235
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   1815
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   4695
      Begin VB.TextBox txtDelString 
         Height          =   270
         Left            =   1560
         TabIndex        =   23
         Text            =   "����,����"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdCalledColor 
         Caption         =   "��"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   400
         Width           =   255
      End
      Begin VB.TextBox txtRect 
         Height          =   345
         Index           =   6
         Left            =   1560
         TabIndex        =   17
         Text            =   "6"
         Top             =   375
         Width           =   735
      End
      Begin VB.TextBox txtRect 
         Height          =   345
         Index           =   5
         Left            =   1560
         TabIndex        =   14
         Text            =   "2"
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "ɾ���ַ���"
         Height          =   255
         Left            =   660
         TabIndex        =   22
         Top             =   1340
         Width           =   975
      End
      Begin VB.Shape shpCalled 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00408000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3480
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "�Ѻ��У�"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   20
         Top             =   885
         Width           =   735
      End
      Begin VB.Shape shpCalling 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3480
         Top             =   400
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "�����У�"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   18
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "���м�¼��ʾ����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   435
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "��"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   885
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "��ѯ���ʱ�䣺"
         Height          =   255
         Left            =   310
         TabIndex        =   13
         Top             =   885
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   3240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "��������"
      Height          =   375
      Left            =   2655
      TabIndex        =   11
      Top             =   4575
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3855
      TabIndex        =   10
      Top             =   4575
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1335
      TabIndex        =   9
      Top             =   4575
      Width           =   1100
   End
   Begin VB.Frame frmRect 
      Caption         =   "Һ����λ�ã��ֱ���Ϊ��λ��"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtRect 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "�߶ȣ�"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "��ȣ�"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "��"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   240
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLCDSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Һ�����Ĳ�����ʹ�ñ���ע�����

Public Function zlShowMe(frmParent As Form) As Boolean
    
    Me.Show 1, frmParent
    
    zlShowMe = True
End Function

Private Sub cmdCalledColor_Click()
    dlgColor.Color = shpCalled.FillColor
    dlgColor.ShowColor
    shpCalled.FillColor = dlgColor.Color
End Sub

Private Sub cmdCancel_Click()
    '�رմ���
    Unload Me
End Sub

Private Sub cmdColor_Click()
    dlgColor.Color = shpCalling.FillColor
    dlgColor.ShowColor
    shpCalling.FillColor = dlgColor.Color
End Sub

Private Sub cmdFont_Click()
    Dim strReg As String
    
    On Error GoTo err
    
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    dlgFont.Flags = cdlCFBoth
    dlgFont.CancelError = False  '�ѵ�ȡ������������
    dlgFont.FontName = GetSetting("ZLSOFT", strReg, "����", "����")
    dlgFont.FontBold = GetSetting("ZLSOFT", strReg, "����", "False")
    dlgFont.FontItalic = GetSetting("ZLSOFT", strReg, "б��", "False")
    dlgFont.FontSize = GetSetting("ZLSOFT", strReg, "�ֺ�", "14")
    dlgFont.ShowFont
    On Error GoTo 0
    '��������
    SaveSetting "ZLSOFT", strReg, "����", dlgFont.FontName
    SaveSetting "ZLSOFT", strReg, "����", dlgFont.FontBold
    SaveSetting "ZLSOFT", strReg, "б��", dlgFont.FontItalic
    SaveSetting "ZLSOFT", strReg, "�ֺ�", dlgFont.FontSize

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    '��Ⲣ�������
    Dim strReg As String
    
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    
    SaveSetting "ZLSOFT", strReg, "��", Val(txtRect(1).Text)
    SaveSetting "ZLSOFT", strReg, "��", Val(txtRect(2).Text)
    SaveSetting "ZLSOFT", strReg, "���", Val(txtRect(3).Text)
    SaveSetting "ZLSOFT", strReg, "�߶�", Val(txtRect(4).Text)
    SaveSetting "ZLSOFT", strReg, "LED��ѯʱ��", Val(txtRect(5).Text)
    SaveSetting "ZLSOFT", strReg, "���м�¼��ʾ��", Val(txtRect(6).Text)
    SaveSetting "ZLSOFT", strReg, "��������ɫ", shpCalling.FillColor
    SaveSetting "ZLSOFT", strReg, "�Ѻ�����ɫ", shpCalled.FillColor
    SaveSetting "ZLSOFT", strReg, "ɾ���ַ�", txtDelString.Text
    
    '�رմ���
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    
    txtRect(1).Text = GetSetting("ZLSOFT", strReg, "��", "1024")
    txtRect(2).Text = GetSetting("ZLSOFT", strReg, "��", "0")
    txtRect(3).Text = GetSetting("ZLSOFT", strReg, "���", "1024")
    txtRect(4).Text = GetSetting("ZLSOFT", strReg, "�߶�", "768")
    txtRect(5).Text = GetSetting("ZLSOFT", strReg, "LED��ѯʱ��", "2")
    txtRect(6).Text = GetSetting("ZLSOFT", strReg, "���м�¼��ʾ��", "6")
    txtDelString.Text = GetSetting("ZLSOFT", strReg, "ɾ���ַ�", "")
    shpCalling.FillColor = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
    shpCalled.FillColor = GetSetting("ZLSOFT", strReg, "�Ѻ�����ɫ", &H408000)
End Sub


Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
