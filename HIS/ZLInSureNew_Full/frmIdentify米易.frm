VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�����޶� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   32
      Top             =   5100
      Width           =   2865
   End
   Begin VB.TextBox txt˳��� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1815
      TabIndex        =   30
      Top             =   4650
      Width           =   2865
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   28
      Top             =   4200
      Width           =   2865
   End
   Begin VB.TextBox txtԭ���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   150
      Width           =   2865
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4290
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�����(M)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5160
      TabIndex        =   36
      Top             =   5010
      Width           =   1470
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5190
      TabIndex        =   35
      Top             =   900
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5190
      TabIndex        =   34
      Top             =   330
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   8115
      Left            =   4920
      TabIndex        =   33
      Top             =   -450
      Width           =   30
   End
   Begin VB.TextBox txt�ʻ���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   26
      Top             =   3750
      Width           =   2865
   End
   Begin VB.TextBox txt��λ���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   24
      Top             =   3300
      Width           =   2865
   End
   Begin VB.TextBox txt��Ա��� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   22
      Top             =   2850
      Width           =   2865
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3870
      TabIndex        =   16
      Top             =   1500
      Width           =   795
   End
   Begin VB.TextBox txt���֤�� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   20
      Top             =   2400
      Width           =   2865
   End
   Begin VB.ComboBox cbo�Ա� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1500
      Width           =   885
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   12
      Top             =   1050
      Width           =   2865
   End
   Begin VB.TextBox txt������ 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   10
      Top             =   600
      Width           =   2865
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2865
   End
   Begin MSMask.MaskEdBox txt�������� 
      Height          =   345
      Left            =   1800
      TabIndex        =   18
      Top             =   1950
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl�����޶� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����޶�(&F)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   31
      Top             =   5160
      Width           =   1425
   End
   Begin VB.Label lbl˳��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���˱��(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   7
      Top             =   1110
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&L)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   345
      TabIndex        =   29
      Top             =   4710
      Width           =   1425
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&Q)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   585
      TabIndex        =   27
      Top             =   4260
      Width           =   1170
   End
   Begin VB.Label lblԭ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   825
      TabIndex        =   0
      Top             =   210
      Width           =   915
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(I)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lbl�ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʻ����(&D)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   25
      Top             =   3810
      Width           =   1425
   End
   Begin VB.Label lbl��λ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��λ����(&W)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   23
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label lbl��Ա��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ա���(&T)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   21
      Top             =   2910
      Width           =   1425
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&G)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2910
      TabIndex        =   15
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lbl���֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���֤��(&K)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   19
      Top             =   2460
      Width           =   1425
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&B)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   17
      Top             =   1995
      Width           =   1425
   End
   Begin VB.Label lbl�Ա� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   11
      Top             =   1110
      Width           =   915
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&R)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   9
      Top             =   660
      Width           =   1425
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&A)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   5
      Top             =   660
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mbytType As Long
Private mstrReturn As String
Private rsTemp As New ADODB.Recordset

Public mstr������ As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    mstr������ = frm�޸�����.ChangePassword(txtԭ����)
End Sub

Private Sub cmdOK_Click()
    Dim strIdentify As String, strAddition As String
    Dim lngResult As Long
    
    If Trim(txt˳���) = "" Or Trim(txt����) = "" Then
        MsgBox "����뿨����������󰴻س���", vbInformation, gstrSysName
        txtԭ����.SetFocus
        Exit Sub
    End If
    
    If Trim(txt������.Text) = "" Then
        MsgBox "δ�õ������ţ��޷�������", vbInformation, gstrSysName
        txtԭ����.SetFocus
        Exit Sub
    End If
    
    '��鲡��״̬
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, type_����, CStr(txt������.Text))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = gComInfo_����.����                         '0����
    strIdentify = strIdentify & ";" & gComInfo_����.���˱��     '1ҽ���ţ����˱�ţ�
    strIdentify = strIdentify & ";" & gComInfo_����.����                              '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & cbo�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & txt��������.Text          '5��������
    strIdentify = strIdentify & ";" & txt���֤��.Text          '6���֤
    strIdentify = strIdentify & ";" & txt��λ����.Text          '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";" & txt������.Text                              '9.˳���
    strAddition = strAddition & ";" & txt��Ա���.Tag           '10��Ա���
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & Val(txt����.Tag)           '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";" & Val(txt����.Text)         '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";0"                            '21סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, type_����)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    
    If mstr������ <> "" Then
        '�����޸�����ӿڣ����ɹ�������ʾ������
        gComInfo_����.������ = mstr������
        gstrPara_���� = "<code>" & gComInfo_����.���� & "</code>" & GetParaCode(����, gComInfo_����.����) & _
            GetParaCode(������, gComInfo_����.������)
        If ���ýӿ�_����("modifypassword") Then
            gComInfo_����.���� = mstr������
            '���¸����ʻ��е���Ϣ
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & type_���� & ",'����','''" & mstr������ & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
            
            'д��:0 as �޸���ȷ��4 as  д������, 3 as ת�����ݴ���,2 as  ��������,1 as ����ԭ�������ȷ
            '�ȵ��õ����׽ӿڣ������ȷִ�У���˵��ԭ����û��
            lngResult = Card_ChangePsd(gintComPort, txtԭ����.Text, mstr������)
            If lngResult <> 0 Then
                gComInfo_����.���� = txtԭ����.Text
                MsgBox "�����޸�ʧ�ܣ��Կɼ���������", vbInformation, gstrSysName
            Else
                gComInfo_����.���� = mstr������
            End If
        End If
    Else
        gComInfo_����.���� = txtԭ����.Text
    End If
    
    Unload Me
End Sub

Private Sub cmd����_Click()
    Dim rs���� As ADODB.Recordset
    
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=[1]"
    Set rs���� = New ADODB.Recordset
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�����֤", type_����)
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(type_����, rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�") = True Then
            txt����.Text = rs����("����")
            txt����.Tag = rs����("ID")
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With cbo�Ա�
        .AddItem "��"
        .AddItem "Ů"
        .ListIndex = 0
    End With
    
    If mbytType = 0 Then
        '����
        Me.Height = 4545
        Me.cmdModify.Top = 3350
    End If
    
    If mlng����ID <> 0 Then Call ReadPatient
End Sub

Private Sub ClearCons()
    '�����������
    txt����.Text = ""
    txt˳���.Text = ""
    txt������ = ""
    txt���� = ""
    cbo�Ա�.ListIndex = 0
    txt���� = ""
    txt�������� = "____-__-__"
    txt���֤�� = ""
    txt��Ա��� = ""
    txt��λ���� = ""
    txt�ʻ���� = ""
    txt���� = ""
    txt�������� = ""
    txt�����޶� = ""
End Sub

Private Sub ReadPatient()
    Dim intField As Integer
    '��ȡ�ò��˵���ϸ��Ϣ
    If mlng����ID = 0 Then Exit Sub
    gstrSQL = " Select A.����,A.ҽ����,A.����,B.����,Decode(B.�Ա�,'Ů',1,0) �Ա�,nvl(B.����,0) ����,Nvl(A.����ID,0) ����ID,C.���� ����," & _
              " A.˳��� ������,to_Char(B.��������,'yyyy-MM-dd') ��������,B.���֤��,B.������λ ��λ����,Nvl(A.�ʻ����,0) �ʻ����,A.��ְ ��Ա���,A.��ע" & _
              " From �����ʻ� A,������Ϣ B,���ղ��� C" & _
              " Where A.����ID=B.����ID And A.����ID=C.ID(+) And A.����ID=[1] And A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ���Ϣ", mlng����ID, type_����)
    
    With rsTemp
        If .EOF Then Exit Sub
        'ֻ����������Ϣ����������Ա�����п����С��޷�����
'        txt���� = !����
'        txt˳��� = !ҽ����
'        txt����.Tag = !����ID
'        txt���� = !����
'        txt������ = NVL(!˳���, "")
        txt���� = Nvl(!����, "")
        cbo�Ա�.ListIndex = !�Ա�
        txt���� = Format(!����, "#####0;#####0; ;")
        txt�������� = !��������
        txt���֤�� = !���֤��
        txt��Ա��� = ת����Ա���(!��Ա���)
        txt��λ���� = Nvl(!��λ����, "")
'        txt�ʻ���� = Format(!�ʻ����, "#####0;#####0; ;")
'        If Not IsNull(!��ע) Then
'            '����;��������;�����޶�
'            txt���� = Format(Split(!��ע, ";")(0), "#####0.00;-#####0.00; ;")
'            txt�������� = Format(Split(!��ע, ";")(1), "#####0.00;-#####0.00; ;")
'            txt�����޶� = Format(Split(!��ע, ";")(2), "#####0.00;-#####0.00; ;")
'        End If
        
'        gComInfo_����.���� = txt����
'        gComInfo_����.���˱�� = txt˳���
'        gComInfo_����.������ = txt������
        gComInfo_����.���� = txt����
        gComInfo_����.�Ա� = Me.cbo�Ա�.ListIndex + 1
        gComInfo_����.���� = txt����
        gComInfo_����.�������� = txt��������
        gComInfo_����.���֤�� = txt���֤��
        gComInfo_����.��λ���� = txt��λ����
'        gComInfo_����.�ʻ���� = Val(txt�ʻ����)
'        gComInfo_����.�������� = Val(txt����)
'        gComInfo_����.����ͳ�ﱨ������ = Val(txt��������)
'        gComInfo_����.����ͳ���޶� = Val(txt�����޶�)
        gComInfo_����.��Ա��� = !��Ա���
'        Call ��ȡ���ֱ���(Val(txt����.Tag))
    End With
End Sub

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0) As String
    mstrReturn = ""
    mlng����ID = lng����ID
    mbytType = bytType
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" Or txt����.Tag <> "" Then
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt����.Text
    gstrSQL = "Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ⲡ','��ͨ��') ��� " & _
             "   FROM ���ղ��� A WHERE A.����=[1] And (A.���� like [2] ||'%' or A.���� like [2] ||'%' or A.���� like [2] ||'%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, type_����, strText)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(type_����, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub WriteFace()
    '������
    txt����.Text = gComInfo_����.����
    txt˳���.Text = gComInfo_����.���˱��
    txt������.Text = gComInfo_����.������
    txt����.Text = gComInfo_����.����
    cbo�Ա�.ListIndex = gComInfo_����.�Ա� - 1
    txt���� = gComInfo_����.����
    txt�������� = gComInfo_����.��������
    txt���֤�� = gComInfo_����.���֤��
    txt��Ա��� = ת����Ա���(gComInfo_����.��Ա���)
    txt��λ���� = gComInfo_����.��λ����
    txt�ʻ���� = Format(gComInfo_����.�ʻ����, "#####0.00;-#####0.00; ;")
    txt���� = gComInfo_����.��������
    txt�������� = gComInfo_����.����ͳ�ﱨ������
    txt�����޶� = gComInfo_����.����ͳ���޶�
End Sub

Private Function ������() As Boolean
    Dim recode As Integer, strResult As String
    strResult = Card_userinfo(gintComPort, txtԭ����.Text, recode)
    If recode <> 0 Then
        MsgBox strResult & "��ȡ�α�������Ϣ���󣬿������������", vbInformation, gstrSysName
        Exit Function
    End If

    gComInfo_����.���˱�� = Mid(strResult, 11, 15)
    gComInfo_����.���� = Mid(strResult, 1, 10)
    gComInfo_����.�ʻ���� = Val(Mid(strResult, 26, 8)) / 100                   '����Ϊ��λ��¼�Ľ��ת��Ϊ��ԪΪ��λ�Ľ��
    gComInfo_����.���� = txtԭ����.Text
    ������ = True
End Function

Private Function ת����Ա���(ByVal str��Ա��� As String) As String
    Select Case str��Ա���
    Case 11, 12
        ת����Ա��� = IIf(str��Ա��� = "11", "��ְ", "��ְ����פ��")
    Case 21, 22
        ת����Ա��� = IIf(str��Ա��� = "21", "����", "������ذ���")
    Case Else
        ת����Ա��� = "����"
    End Select
End Function

Private Sub txtԭ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtԭ����.Text) = "" Then
        MsgBox "���������룡", vbInformation, gstrSysName
        txtԭ����.SetFocus
        Exit Sub
    End If
    
    '������沢����
    Call ClearCons
    If Not ������ Then Exit Sub
    If Trim(gComInfo_����.���˱��) <> "" And Trim(gComInfo_����.����) <> "" Then
        '���������֤�ӿ�
        If mbytType = 0 Then     '������������֤�ӿ�
            gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
                GetParaCode(����������, gComInfo_����.����������)
            If Not ���ýӿ�_����("identifyinfogetting") Then Exit Sub
        Else                    'סԺ���ô��������ӿ�
            If Not ���ýӿ�_����("getsysdate") Then Exit Sub
            
            gComInfo_����.֧����� = "0301"    '��ͨסԺ
            Call ��ȡ���ֱ���(Val(txt����.Tag))
            
            gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
                GetParaCode(����, gComInfo_����.����) & GetParaCode(����������, gComInfo_����.����������) & _
                GetParaCode(֧�����, gComInfo_����.֧�����) & GetParaCode(���ֱ���, gComInfo_����.���ֱ���) & _
                GetParaCode(������ȡʱ��, gComInfo_����.ϵͳʱ��)
            If Not ���ýӿ�_����("audittreatment") Then Exit Sub
        End If
        Call WriteFace
    End If
End Sub
