VERSION 5.00
Begin VB.Form frmAppforCritical 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Σ��ֵ����"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   Icon            =   "frmAppforCritical.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6435
      Left            =   30
      ScaleHeight     =   6405
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   15
      Width           =   12825
      Begin VB.CommandButton cmdSend 
         Caption         =   "����"
         Height          =   375
         Left            =   9300
         TabIndex        =   38
         Top             =   5820
         Width           =   1305
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�"
         Height          =   375
         Left            =   10980
         TabIndex        =   37
         Top             =   5820
         Width           =   1305
      End
      Begin VB.TextBox txtReturn 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   1290
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3840
         Width           =   11310
      End
      Begin VB.TextBox txtIPTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtStartSampleNO 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox TxtBad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtAppforDoctor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox txtVerifyDoctor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   645
         Width           =   1815
      End
      Begin VB.TextBox txtStartAge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3345
         Width           =   11310
      End
      Begin VB.TextBox TxtHealthNo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtSimpleTyppe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   4470
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtPatientType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox txtIPName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtSay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2835
         Width           =   11310
      End
      Begin VB.TextBox txtVerifyTime 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʩ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   36
         Top             =   3840
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֪ͨʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   34
         Top             =   2295
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   13
         X1              =   10740
         X2              =   12570
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   32
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   31
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   2295
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   27
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6540
         TabIndex        =   26
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9690
         TabIndex        =   25
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�걾����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   24
         Top             =   2295
         Width           =   960
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   165
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6540
         TabIndex        =   22
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   21
         Top             =   3330
         Width           =   960
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ٴ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   20
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label lblinto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Σ��ֵ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   570
         Index           =   0
         Left            =   5355
         TabIndex        =   19
         Top             =   90
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   18
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֪ͨ��ʦ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6540
         TabIndex        =   17
         Top             =   2295
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   0
         X1              =   1260
         X2              =   3090
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   3
         X1              =   4470
         X2              =   6300
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   1
         X1              =   1260
         X2              =   3090
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   2
         X1              =   7650
         X2              =   9480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   4
         X1              =   10740
         X2              =   12570
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   5
         X1              =   1260
         X2              =   3090
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   6
         X1              =   10740
         X2              =   12570
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   7
         X1              =   7650
         X2              =   9480
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   8
         X1              =   4470
         X2              =   6300
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   9
         X1              =   1260
         X2              =   3090
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   10
         X1              =   4470
         X2              =   6300
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   11
         X1              =   7650
         X2              =   9480
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   12
         X1              =   1260
         X2              =   12540
         Y1              =   3645
         Y2              =   3645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֪ͨ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   2835
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   15
         X1              =   1260
         X2              =   12600
         Y1              =   3135
         Y2              =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Index           =   16
         X1              =   10740
         X2              =   12570
         Y1              =   945
         Y2              =   945
      End
   End
End
Attribute VB_Name = "frmAppforCritical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSend As Boolean         '�Ƿ��ͳɹ�
Private mstrDicName As String       '������ʦ
Private mstrReturn As String         '����ҽ����д��Ϣ

Public Function ShowMe(objFrm As Object, ByVal strDicName As String, ByVal lngSampleID As Long, Optional ByRef strReturn As String) As Boolean
    mstrReturn = ""
    mstrDicName = strDicName
    Call InitData(lngSampleID)
    Me.Show 1, objFrm
    '�������ʧ��,��strReturn���ؿ�
    If mblnSend = False Then
        strReturn = ""
    Else
        strReturn = mstrReturn
    End If
    ShowMe = mblnSend
End Function

Private Sub InitData(ByVal lngSampleID As Long, Optional strErr As String)
          '��ʼ������
          'lngSampleID        �걾ID
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo InitData_Error

2         strSQL = "Select Distinct a.ID �걾ID, a.�걾���, a.������, a.�������, a.����, a.�Ա�, a.����," & _
                  "a.������Դ, a.����ʱ�� ����ʱ��, a.�걾����, a.����, a.������, a.����ʱ��," & _
                  "a.������ ����ҽ��,B.֪ͨ���� , b.֪ͨ�� ֪ͨ��ʦ,b.֪ͨʱ��, b.��ע ��ʦ��ע" & _
                  " From ���鱨���¼ A, ����Σ��ֵ��¼ B" & _
                  " Where  a.Id = b.�걾id and a.Id =[1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "Σ��֪ͨ��ѯ", lngSampleID)
4         If rsTmp.RecordCount > 0 Then
5             txtVerifyDoctor = rsTmp("������") & ""
6             txtVerifyTime = rsTmp("����ʱ��") & ""
7             txtStartSampleNO = rsTmp("�걾���") & ""
8             txtStartSampleNO.Tag = rsTmp("�걾ID") & ""
9             txtName = rsTmp("����") & ""
10            If Val("" & rsTmp!�Ա�) = 1 Then
11                txtSex = "��"
12            ElseIf Val("" & rsTmp!�Ա�) = 2 Then
13               txtSex = "Ů"
14            Else
15                txtSex = "δ֪"
16            End If
17            txtStartAge = rsTmp("����") & ""
18            txtID = rsTmp("������") & ""
19            TxtHealthNo = rsTmp("�������") & ""
20            txtAppforDoctor = rsTmp("����ҽ��") & ""
21            txtPatientType = rsTmp("����ʱ��") & ""
22            TxtBad = rsTmp("����") & ""
23            txtSimpleTyppe = rsTmp("�걾����") & ""
24            txtSay = rsTmp("֪ͨ����") & ""
25            txtIPName = rsTmp("֪ͨ��ʦ") & ""
26            txtIPTime = rsTmp("֪ͨʱ��") & ""
27            txtRemark = rsTmp("��ʦ��ע") & ""
28        End If

29        rsTmp.Close
30        Set rsTmp = Nothing
          
31        Exit Sub
InitData_Error:
32        mblnSend = False
33        Call WriteErrLog("zlPublicHisCommLis", "frmAppforCritical", "ִ��(InitData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
34        Err.Clear
End Sub

Private Sub cmdExit_Click()
    mblnSend = False
    Unload Me
End Sub

Private Sub cmdSend_Click()
          '���ͷ�����LIS
          Dim strSQL As String
          Dim strTime As String
          Dim lngSampleID As String
          Dim strInfo As String           'ҽ����д�����ʩ
              
1         On Error GoTo cmdSend_Click_Error

2         lngSampleID = Val(Trim(txtStartSampleNO.Tag))           '��ȡ�걾��
3         strTime = Format(Currentdate, "yyyy-mm-dd hh:mm:ss")    '������ʱ��
4         strInfo = Trim(Me.txtReturn.Text)
              
5         gcnLisOracle.BeginTrans
           '������Ϣ����
6         strSQL = "Zl_������Ϣ��¼_Edit(1,2,null," & lngSampleID & ",null,null,null,'ҽ���Ѳ���Σ��ֵ','Σ��ֵ')"
7         Call ComExecuteProc(Sel_Lis_DB, strSQL, "������Ϣ��¼")
          
8         strSQL = "Zl_����Σ��ֵ��¼_Message(" & lngSampleID & ",'" & mstrDicName & "',to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),'" & strInfo & "')"
9         Call ComExecuteProc(Sel_Lis_DB, strSQL, "Σ��ֵ��¼֪ͨ")
10        gcnLisOracle.CommitTrans
          
11        SaveDBLog 18, 6, Val(lngSampleID), "Σ��ֵ����", "ȷ�ϴ���ȷ���ˣ�" & mstrDicName & " ȷ��ʱ��:" & Format(strTime, "yyyy-MM-dd HH:mm:ss") & " �����ʩ:" & strInfo, 2500, "�ٴ�ʵ���ҹ���"
          
12        If strInfo <> "" Then mstrReturn = strInfo
          
13        mblnSend = True
14        MsgBox "���ͳɹ�", vbInformation, Me.Caption
15        Unload Me

16        Exit Sub
cmdSend_Click_Error:
17        gcnLisOracle.RollbackTrans
18        mblnSend = False
19        Call WriteErrLog("zlPublicHisCommLis", "frmAppforCritical", "ִ��(cmdSend_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrDicName = ""
End Sub
