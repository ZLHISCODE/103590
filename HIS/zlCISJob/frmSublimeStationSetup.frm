VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSublimeStationSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   Icon            =   "frmSublimeStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picControl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   2400
      ScaleHeight     =   1785
      ScaleWidth      =   2295
      TabIndex        =   37
      Top             =   3450
      Visible         =   0   'False
      Width           =   2295
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1497
         Width           =   200
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   60
         Picture         =   "frmSublimeStationSetup.frx":000C
         ScaleHeight     =   1350
         ScaleWidth      =   2160
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   90
         Width           =   2160
         Begin VB.Shape shpBorder 
            BorderColor     =   &H00C56A31&
            FillColor       =   &H00FF8080&
            Height          =   270
            Left            =   1890
            Top             =   1080
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Shape shpValue 
            BorderColor     =   &H00C56A31&
            FillColor       =   &H00FF8080&
            Height          =   270
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   270
         End
      End
      Begin VB.Label lblColor 
         Caption         =   "&HFFFFFF"
         Height          =   195
         Left            =   405
         TabIndex        =   41
         Top             =   1500
         UseMnemonic     =   0   'False
         Width           =   1365
      End
   End
   Begin VB.CheckBox chkBedInfo 
      Caption         =   "���������������λ������ʾ��λʹ��״��"
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   5685
      Width           =   3900
   End
   Begin VB.Frame fraFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   3
      Left            =   1740
      TabIndex        =   43
      Top             =   5580
      Width           =   300
   End
   Begin VB.TextBox txt��Ժ���� 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   1725
      MaxLength       =   2
      TabIndex        =   32
      Text            =   "3"
      Top             =   5400
      Width           =   300
   End
   Begin MSComctlLib.ImageList img24 
      Left            =   210
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraFilter 
      Caption         =   " ����ȼ���ɫ"
      Height          =   1530
      Index           =   1
      Left            =   180
      TabIndex        =   28
      Top             =   3735
      Width           =   4590
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   3
         Left            =   3840
         Picture         =   "frmSublimeStationSetup.frx":0782
         Stretch         =   -1  'True
         Top             =   900
         Width           =   345
      End
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   2
         Left            =   1770
         Picture         =   "frmSublimeStationSetup.frx":0E84
         Stretch         =   -1  'True
         Top             =   900
         Width           =   345
      End
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   3840
         Picture         =   "frmSublimeStationSetup.frx":1586
         Stretch         =   -1  'True
         Top             =   420
         Width           =   345
      End
      Begin VB.Image img����ȼ� 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1770
         Picture         =   "frmSublimeStationSetup.frx":1C88
         Stretch         =   -1  'True
         Top             =   420
         Width           =   345
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2610
         TabIndex        =   31
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   540
         TabIndex        =   30
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2610
         TabIndex        =   29
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lbl����ȼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ؼ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   21
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   2340
      TabIndex        =   24
      Top             =   6015
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   3555
      TabIndex        =   25
      Top             =   6015
      Width           =   1100
   End
   Begin VB.Frame fraAdvice 
      Caption         =   " ҽ���������� "
      Height          =   2790
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   4590
      Begin VB.CheckBox chkNurse 
         Caption         =   "�������廤����Ϣ����"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   2535
         Width           =   2235
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Ѫ������"
         Height          =   195
         Index           =   12
         Left            =   1560
         TabIndex        =   44
         Top             =   1680
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   11
         Left            =   465
         TabIndex        =   15
         Top             =   1680
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Һ�ܾ�"
         Height          =   195
         Index           =   5
         Left            =   1350
         TabIndex        =   9
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Σ��ֵ"
         Height          =   195
         Index           =   4
         Left            =   465
         TabIndex        =   8
         Top             =   1125
         Width           =   870
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RISԤԼ׼��"
         Height          =   195
         Index           =   8
         Left            =   465
         TabIndex        =   12
         Top             =   1410
         Width           =   1320
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "ȡѪ֪ͨ"
         Height          =   195
         Index           =   9
         Left            =   1845
         TabIndex        =   13
         Top             =   1410
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�걾����"
         Height          =   195
         Index           =   10
         Left            =   3015
         TabIndex        =   14
         Top             =   1410
         Width           =   1025
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RISԤԼ"
         Height          =   195
         Index           =   7
         Left            =   3525
         TabIndex        =   11
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   2205
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   1770
         TabIndex        =   18
         Top             =   2130
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��������"
         Height          =   195
         Index           =   6
         Left            =   2460
         TabIndex        =   10
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox ChkCollate 
         Caption         =   "ҽ��������Զ���λ������ҽ��ҳ��"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1920
         Width           =   3900
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "����"
         Height          =   195
         Index           =   3
         Left            =   3270
         TabIndex        =   7
         Top             =   885
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�¿�"
         Height          =   195
         Index           =   0
         Left            =   1185
         TabIndex        =   4
         Top             =   885
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��ͣ"
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   5
         Top             =   885
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�·�"
         Height          =   195
         Index           =   2
         Left            =   2580
         TabIndex        =   6
         Top             =   885
         Width           =   660
      End
      Begin VB.TextBox txtNotifyAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "10"
         Top             =   315
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   34
         Top             =   495
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdviceDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   33
         Top             =   765
         Width           =   300
      End
      Begin VB.TextBox txtNotifyAdviceDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "1"
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkNotifyAdvice 
         Caption         =   "ÿ    �����Զ�ˢ��ҽ�����������е�����"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   300
         TabIndex        =   36
         Top             =   885
         Width           =   810
      End
      Begin VB.Label lblNotifyAdviceDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ���ڴ����ҽ��������ʾ����������"
         Height          =   180
         Left            =   570
         TabIndex        =   35
         Top             =   600
         Width           =   3420
      End
   End
   Begin VB.Frame fraFilter 
      Caption         =   " ���Ի��������� "
      Height          =   690
      Index           =   0
      Left            =   180
      TabIndex        =   26
      Top             =   2940
      Width           =   4590
      Begin VB.Frame fraFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   2
         Left            =   1320
         TabIndex        =   42
         Top             =   495
         Width           =   300
      End
      Begin VB.TextBox txt������� 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "3"
         Top             =   315
         Width           =   300
      End
      Begin VB.CheckBox chkPatientFilter 
         Caption         =   "��ȡ���    ���ڵ�סԺ����"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   315
         Width           =   3900
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   180
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   40
      Top             =   3390
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CheckBox chkNewPati 
      Caption         =   "������б���ʾ    ���ڵǼǵ�סԺ����"
      Height          =   195
      Left            =   180
      TabIndex        =   22
      Top             =   5400
      Width           =   3900
   End
End
Attribute VB_Name = "frmSublimeStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarColor As OLE_COLOR
Public mstrPrivs As String
Private mlngModual As Long

Private Const ALTERNATE = 1
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" _
    (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Polyline Lib "gdi32" _
    (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'�趨һ�����岶����꣬���������������Ϣ�������ô���
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private mlngColor As Long
Private mintIndex As Long
Private mobjFileSys As New FileSystemObject

Public Sub ShowMe()
    '���°�סԺ��ʿ����վ���ã���ʾ��ע��ť
    mintIndex = 0
    Me.Show vbModal
End Sub

Private Sub chkNewPati_Click()
    On Error Resume Next
    If chkNewPati.Value = 1 Then
        txt��Ժ����.Enabled = True
        txt��Ժ����.SetFocus
    Else
        txt��Ժ����.Enabled = False
        txt��Ժ����.Text = ""
    End If
End Sub

Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkPatientFilter_Click()
    On Error Resume Next
    If chkPatientFilter.Value = 1 Then
        txt�������.Enabled = True
        txt�������.SetFocus
    Else
        txt�������.Enabled = False
        txt�������.Text = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
        If txtNotifyAdvice.Text = "" Then
            MsgBox "������ҽ�����ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
        Else
            MsgBox "ҽ�����ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
        End If
        txtNotifyAdvice.SetFocus: Exit Sub
    End If
    If Val(txtNotifyAdviceDay.Text) = 0 Then
        If txtNotifyAdviceDay.Text = "" Then
            MsgBox "������Ҫ���ѵ�ҽ��������", vbInformation, gstrSysName
        Else
            MsgBox "Ҫ���ѵ�ҽ����������ӦΪ1�졣", vbInformation, gstrSysName
        End If
        txtNotifyAdviceDay.SetFocus: Exit Sub
    End If
    If chkPatientFilter.Value = 1 Then
        If Trim(txt�������.Text) = "" Then
            MsgBox "�������������������", vbInformation, gstrSysName
            txt�������.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt�������.Text) Then
            MsgBox "��������к��зǷ��ַ�����ֻ���������֣�", vbInformation, gstrSysName
            txt�������.SetFocus
            Exit Sub
        End If
        If Val(txt�������.Text) <= 0 Then
            MsgBox "���������������㣡", vbInformation, gstrSysName
            txt�������.SetFocus
            Exit Sub
        End If
    End If
    
    If chkNewPati.Value = 1 Then
        If Trim(txt��Ժ����.Text) = "" Then
            MsgBox "������������ʾ����Ժ�Ǽ�������", vbInformation, gstrSysName
            txt��Ժ����.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txt��Ժ����.Text) Then
            MsgBox "�������ʾ����Ժ�Ǽ������к��зǷ��ַ�����ֻ���������֣�", vbInformation, gstrSysName
            txt��Ժ����.SetFocus
            Exit Sub
        End If
        If Val(txt��Ժ����.Text) <= 0 Then
            MsgBox "�������ʾ����Ժ�Ǽ�������������㣡", vbInformation, gstrSysName
            txt��Ժ����.SetFocus
            Exit Sub
        End If
    End If
    
    '�Զ�ˢ��ҽ������
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zldatabase.SetPara("�Զ�ˢ��ҽ�����", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, pסԺ��ʿվ, blnSetup)
    Call zldatabase.SetPara("�Զ�ˢ��ҽ������", Val(txtNotifyAdviceDay.Text), glngSys, pסԺ��ʿվ, blnSetup)
    Call zldatabase.SetPara("����������ʾ", chkSound.Value, glngSys, pסԺ��ʿվ, blnSetup)
    strTmp = ""
    For i = 0 To chkWarn.UBound
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zldatabase.SetPara("�Զ�ˢ��ҽ������", strTmp, glngSys, pסԺ��ʿվ, blnSetup)
    
    '�����������
    If chkPatientFilter.Value = 1 Then
        Call zldatabase.SetPara("�������", txt�������.Text, glngSys, 1265, blnSetup)
    Else
        Call zldatabase.SetPara("�������", "0", glngSys, 1265, blnSetup)
    End If
    '������Ժ���� 111016
    If chkNewPati.Value = 1 Then
        Call zldatabase.SetPara("��Ժ����", txt��Ժ����.Text, glngSys, 1265, blnSetup)
    Else
        Call zldatabase.SetPara("��Ժ����", "0", glngSys, 1265, blnSetup)
    End If
    
    '���滤��ȼ�����ɫ
    Call zldatabase.SetPara("�ؼ�������ɫ", img����ȼ�(0).Tag, glngSys, 1265, blnSetup)
    Call zldatabase.SetPara("һ��������ɫ", img����ȼ�(1).Tag, glngSys, 1265, blnSetup)
    Call zldatabase.SetPara("����������ɫ", img����ȼ�(2).Tag, glngSys, 1265, blnSetup)
    Call zldatabase.SetPara("����������ɫ", img����ȼ�(3).Tag, glngSys, 1265, blnSetup)
    '54370:������,2013-05-02,��Ӳ���"ҽ��������Զ���λ��ҽ��ҳ��"
    Call zldatabase.SetPara("ҽ��������Զ���λ��ҽ��ҳ��", ChkCollate.Value, glngSys, 1265, blnSetup)
    Call zldatabase.SetPara("����λ������ʾ��λ״��", chkBedInfo.Value, glngSys, 1265, blnSetup)
    '132721:������,2018-11-17,��Ӳ���"��ʾ���廤����Ϣ"
    Call zldatabase.SetPara("��ʾ���廤����Ϣ", chkNurse.Value, glngSys, 1265, blnSetup And gbln�������廤��ӿ�)
    gblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        ReleaseCapture
        picControl.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strPar As String
    Dim intType As Integer
    
    gblnOK = False
    mlngModual = pסԺ��ʿվ
            
    chkWarn(9).Visible = gblnѪ��ϵͳ
    chkWarn(11).Visible = gblnѪ��ϵͳ
    '�Զ�ˢ��ҽ������
    strPar = zldatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "��������") > 0, intType)
    If Val(strPar) > 0 Then
        chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
    End If
   
    'ǰ���¼��л��Զ����ã���˺���ǿ������
    If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
        txtNotifyAdvice.Enabled = False
    End If
    
    strPar = zldatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "��������") > 0)
    txtNotifyAdviceDay.Text = Val(strPar)
    
    strPar = zldatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, "0000000000000", Array(lbl��������, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "��������") > 0)
    For i = 1 To Len(strPar)
        If i - 1 <= chkWarn.UBound Then
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        End If
    Next
    txt�������.Text = zldatabase.GetPara("�������", glngSys, 1265, "3", Array(chkPatientFilter, txt�������))
    chkPatientFilter.Value = IIf(Val(txt�������.Text) = 0, 0, 1)
    txt�������.Enabled = (chkPatientFilter.Value = 1)
    '111016
    txt��Ժ����.Text = zldatabase.GetPara("��Ժ����", glngSys, 1265, "0", Array(chkNewPati, txt��Ժ����))
    chkNewPati.Value = IIf(Val(txt��Ժ����.Text) = 0, 0, 1)
    txt��Ժ����.Enabled = (chkNewPati.Value = 1)
    '54370:������,2013-05-02,��Ӳ���"ҽ��������Զ���λ��ҽ��ҳ��"
    strPar = zldatabase.GetPara("ҽ��������Զ���λ��ҽ��ҳ��", glngSys, 1265, 0, Array(ChkCollate), InStr(mstrPrivs, "��������") > 0)
    ChkCollate.Value = IIf(Val(strPar) = 1, 1, 0)
    strPar = zldatabase.GetPara("����������ʾ", glngSys, mlngModual, 0, Array(chkSound, cmdSoundSet), InStr(mstrPrivs, "��������") > 0)
    chkSound.Value = IIf(Val(strPar) = 1, 1, 0)
    strPar = zldatabase.GetPara("����λ������ʾ��λ״��", glngSys, 1265, 1, Array(chkBedInfo), InStr(mstrPrivs, "��������") > 0)
    chkBedInfo.Value = IIf(Val(strPar) = 1, 1, 0)
    '132721:������,2018-11-17,��Ӳ���"��ʾ���廤����Ϣ"
    strPar = zldatabase.GetPara("��ʾ���廤����Ϣ", glngSys, 1265, 0, Array(chkNurse), InStr(mstrPrivs, "��������") > 0 And gbln�������廤��ӿ�)
    chkNurse.Value = IIf(Val(strPar) = 1, 1, 0)
    If chkNurse.Enabled = True Then chkNurse.Enabled = gbln�������廤��ӿ�
    chkNurse.Visible = gbln�������廤��ӿ�
    Call InitColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DeleteFile
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng�ؼ� As Long, lngһ�� As Long, lng���� As Long, lng���� As Long
    Const c��ɫ As Long = 8388736
    Const c��ɫ As Long = 255
    Const c��ɫ As Long = 16711680
    Const c��ɫ As Long = 16777215
    
    Call DeleteFile
    '��ȡ����ȼ���������(����ȡȱʡ����)
    strValue = zldatabase.GetPara("�ؼ�������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(0)))
    lng�ؼ� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zldatabase.GetPara("һ��������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(1)))
    lngһ�� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zldatabase.GetPara("����������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(2)))
    lng���� = IIf(strValue = "", c��ɫ, Val(strValue))
    strValue = zldatabase.GetPara("����������ɫ", glngSys, 1265, "", Array(lbl����ȼ�(3)))
    lng���� = IIf(strValue = "", c��ɫ, Val(strValue))
    
    '��ͼ
    mlngColor = lng�ؼ�
    Call DrawPoly
    img����ȼ�(0).Tag = mlngColor
    img����ȼ�(0).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lngһ��
    Call DrawPoly
    img����ȼ�(1).Tag = mlngColor
    img����ȼ�(1).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng����
    Call DrawPoly
    img����ȼ�(2).Tag = mlngColor
    img����ȼ�(2).Picture = img24.ListImages("K_" & mintIndex).Picture
    mlngColor = lng����
    Call DrawPoly
    img����ȼ�(3).Tag = mlngColor
    img����ȼ�(3).Picture = img24.ListImages("K_" & mintIndex).Picture
End Sub

Private Sub img����ȼ�_Click(Index As Integer)
    picControl.Tag = Index
    picControl.Visible = True
    Call SetCOLOR(Val(img����ȼ�(Index).Tag))
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < Picture1.ScaleWidth And Y > 0 And Y < Picture1.ScaleHeight Then
        SetCapture Picture1.hwnd
        shpBorder.Visible = True
    Else
        ReleaseCapture
        shpBorder.Visible = False
    End If

    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    shpBorder.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    
    If Picture1.Point(lX, lY) = -1 Then Exit Sub
    picColor.BackColor = Picture1.Point(lX, lY)
    Select Case CStr(Hex(picColor.BackColor))
    Case "0"
        lblColor = "��ɫ"
    Case "3399"
        lblColor = "��ɫ"
    Case "3333"
        lblColor = "���ɫ"
    Case "3300"
        lblColor = "����"
    Case "663300"
        lblColor = "����"
    Case "800000"
        lblColor = "����"
    Case "993333"
        lblColor = "����"
    Case "333333"
        lblColor = "��ɫ-80%"
    Case "80"
        lblColor = "���"
    Case "66FF"
        lblColor = "��ɫ"
    Case "8080"
        lblColor = "���"
    Case "8000"
        lblColor = "��ɫ"
    Case "808000"
        lblColor = "��ɫ"
    Case "FF0000"
        lblColor = "��ɫ"
    Case "996666"
        lblColor = "��-��"
    Case "808080"
        lblColor = "��ɫ-50%"
    Case "FF"
        lblColor = "��ɫ"
    Case "99FF"
        lblColor = "ǳ��ɫ"
    Case "CC99"
        lblColor = "���ɫ"
    Case "669933"
        lblColor = "����"
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
    Case "FF6633"
        lblColor = "ǳ��"
    Case "800080"
        lblColor = "������"
    Case "999999"
        lblColor = "��ɫ-40%"
    Case "FF00FF"
        lblColor = "�ۺ�"
    Case "CCFF"
        lblColor = "��ɫ"
    Case "FFFF"
        lblColor = "��ɫ"
    Case "FF00"
        lblColor = "����"
    Case "FFFF00"
        lblColor = "����"
    Case "FFCC00"
        lblColor = "����"
    Case "663399"
        lblColor = "÷��"
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
    Case "CC99FF"
        lblColor = "õ���"
    Case "99CCFF"
        lblColor = "��ɫ"
    Case "99FFFF"
        lblColor = "ǳ��"
    Case "CCFFCC"
        lblColor = "ǳ��"
    Case "FFFFCC"
        lblColor = "ǳ����"
    Case "FFCC99"
        lblColor = "����"
    Case "FF99CC"
        lblColor = "����"
    Case "FFFFFF"
        lblColor = "��ɫ"
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lX As Long, lY As Long
    lRow = Y \ (18 * Screen.TwipsPerPixelY)
    lCol = X \ (18 * Screen.TwipsPerPixelX)
    lX = ((lCol) * 18 + 4) * Screen.TwipsPerPixelX
    lY = ((lRow) * 18 + 4) * Screen.TwipsPerPixelY
    picControl.Visible = False
    
    '��ָ����ɫ��ͼ
    mlngColor = picColor.BackColor
    img����ȼ�(Val(picControl.Tag)).Tag = mlngColor
    Call DrawPoly
    img����ȼ�(Val(picControl.Tag)).Picture = img24.ListImages("K_" & mintIndex).Picture
End Sub


Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub SetCOLOR(vData As OLE_COLOR)
    mvarColor = vData
    Dim lRow As Long, lCol As Long
    shpValue.Visible = True
    Select Case CStr(Hex(vData))
    Case "0"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 0
    Case "3399"
        lblColor = "��ɫ"
        lRow = 0
        lCol = 1
    Case "3333"
        lblColor = "���ɫ"
        lRow = 0
        lCol = 2
    Case "3300"
        lblColor = "����"
        lRow = 0
        lCol = 3
    Case "663300"
        lblColor = "����"
        lRow = 0
        lCol = 4
    Case "800000"
        lblColor = "����"
        lRow = 0
        lCol = 5
    Case "993333"
        lblColor = "����"
        lRow = 0
        lCol = 6
    Case "333333"
        lblColor = "��ɫ-80%"
        lRow = 0
        lCol = 7
    Case "80"
        lblColor = "���"
        lRow = 1
        lCol = 0
    Case "66FF"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 1
    Case "8080"
        lblColor = "���"
        lRow = 1
        lCol = 2
    Case "8000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 3
    Case "808000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 4
    Case "FF0000"
        lblColor = "��ɫ"
        lRow = 1
        lCol = 5
    Case "996666"
        lblColor = "��-��"
        lRow = 1
        lCol = 6
    Case "808080"
        lblColor = "��ɫ-50%"
        lRow = 1
        lCol = 7
    Case "FF"
        lblColor = "��ɫ"
        lRow = 2
        lCol = 0
    Case "99FF"
        lblColor = "ǳ��ɫ"
        lRow = 2
        lCol = 1
    Case "CC99"
        lblColor = "���ɫ"
        lRow = 2
        lCol = 2
    Case "669933"
        lblColor = "����"
        lRow = 2
        lCol = 3
    Case "CCCC33"
        lblColor = "ˮ��ɫ"
        lRow = 2
        lCol = 4
    Case "FF6633"
        lblColor = "ǳ��"
        lRow = 2
        lCol = 5
    Case "800080"
        lblColor = "������"
        lRow = 2
        lCol = 6
    Case "999999"
        lblColor = "��ɫ-40%"
        lRow = 2
        lCol = 7
    Case "FF00FF"
        lblColor = "�ۺ�"
        lRow = 3
        lCol = 0
    Case "CCFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 1
    Case "FFFF"
        lblColor = "��ɫ"
        lRow = 3
        lCol = 2
    Case "FF00"
        lblColor = "����"
        lRow = 3
        lCol = 3
    Case "FFFF00"
        lblColor = "����"
        lRow = 3
        lCol = 4
    Case "FFCC00"
        lblColor = "����"
        lRow = 3
        lCol = 5
    Case "663399"
        lblColor = "÷��"
        lRow = 3
        lCol = 6
    Case "C0C0C0"
        lblColor = "��ɫ-25%"
        lRow = 3
        lCol = 7
    Case "CC99FF"
        lblColor = "õ���"
        lRow = 4
        lCol = 0
    Case "99CCFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 1
    Case "99FFFF"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 2
    Case "CCFFCC"
        lblColor = "ǳ��"
        lRow = 4
        lCol = 3
    Case "FFFFCC"
        lblColor = "ǳ����"
        lRow = 4
        lCol = 4
    Case "FFCC99"
        lblColor = "����"
        lRow = 4
        lCol = 5
    Case "FF99CC"
        lblColor = "����"
        lRow = 4
        lCol = 6
    Case "FFFFFF"
        lblColor = "��ɫ"
        lRow = 4
        lCol = 7
    Case Else
        lblColor = "&H" & CStr(Hex(picColor.BackColor))
    End Select
    shpBorder.Visible = False
    shpValue.Move lCol * 18 * Screen.TwipsPerPixelX, lRow * 18 * Screen.TwipsPerPixelY, 270, 270
    shpValue.Visible = True
    If vData = tomAutoColor Or vData = -1 Then
    
    Else
        picColor.BackColor = vData
    End If
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '������Ϊ�ļ�,���������ͼƬʱ,���뵽imagelist���ʼ��ֻ�����һ��,Ӧ��������image�б������ͼƬID���
    
    strFile = App.Path & "\HLDJTMP" & mintIndex & ".BMP"
    SavePicture PicDraw.Image, strFile
    PicDraw.Picture = LoadPicture(strFile)
    img24.ListImages.Add , "K_" & mintIndex, PicDraw.Picture
End Sub

Private Sub DrawPoly()
    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim PtInPoly() As POINTAPI

    '������򲢻�����
    ReDim PtInPoly(4) As POINTAPI
    PtInPoly(1).X = 0
    PtInPoly(1).Y = 0
    PtInPoly(2).X = PicDraw.ScaleWidth
    PtInPoly(2).Y = 0
    PtInPoly(3).X = PicDraw.ScaleWidth
    PtInPoly(3).Y = PicDraw.ScaleHeight
    PtInPoly(4).X = PtInPoly(1).X
    PtInPoly(4).Y = PtInPoly(1).Y
    
    '����ϵͳˢ��
    PicDraw.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '�������ˢ�ӳɹ�,��ѡ��
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn PicDraw.hDC, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    End If
    PicDraw.Refresh
    
    Call AddColor
End Sub

Private Sub DeleteFile()
    Dim objFile As File
    For Each objFile In mobjFileSys.GetFolder(App.Path).Files
        If Left(objFile.Name, 7) = "HLDJTMP" Then
            mobjFileSys.DeleteFile objFile.Path, True
        End If
    Next
End Sub

Private Sub txt�������_GotFocus()
    txt�������.SelStart = 0
    txt�������.SelLength = txt�������.MaxLength
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ժ����_GotFocus()
    txt��Ժ����.SelStart = 0
    txt��Ժ����.SelLength = txt��Ժ����.MaxLength
End Sub

Private Sub txt��Ժ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
