VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���ֽ���༭ 
   Caption         =   "���ֽ���༭"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   FillColor       =   &H000000FF&
   Icon            =   "frm���ֽ���༭.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11520
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdViewArchive 
      Caption         =   "���Ӳ���(&V)"
      Height          =   350
      Left            =   1305
      TabIndex        =   43
      Top             =   7380
      Width           =   1155
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   42
      Top             =   7380
      Width           =   1100
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "�Զ�(&A)"
      Height          =   350
      Left            =   7260
      TabIndex        =   41
      Top             =   7380
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "��ֹ(&S)"
      Height          =   350
      Left            =   9750
      TabIndex        =   40
      Top             =   7935
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.PictureBox picLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7080
      Left            =   0
      ScaleHeight     =   7020
      ScaleWidth      =   3060
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   3120
      Begin VB.PictureBox pic��Ŀ��Ϣ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2820
         Left            =   135
         ScaleHeight     =   2820
         ScaleWidth      =   2790
         TabIndex        =   28
         Top             =   4050
         Width           =   2790
         Begin VB.PictureBox imgXMXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2460
            Picture         =   "frm���ֽ���༭.frx":000C
            ScaleHeight     =   210
            ScaleWidth      =   255
            TabIndex        =   39
            Top             =   80
            Width           =   255
         End
         Begin VB.TextBox txt��Ŀ��Ϣ 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   450
            Width           =   2490
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "��Ŀ��Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   29
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox pic������Ϣ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   135
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   21
         Top             =   2220
         Width           =   2790
         Begin VB.PictureBox imgFAXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2460
            Picture         =   "frm���ֽ���༭.frx":005B
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   38
            Top             =   80
            Width           =   255
         End
         Begin VB.Label lbl���� 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   195
            Left            =   225
            TabIndex        =   27
            Top             =   682
            Width           =   2580
         End
         Begin VB.Label lbl�ܷ� 
            BackStyle       =   0  'Transparent
            Caption         =   "�ܷ�:"
            Height          =   195
            Left            =   225
            TabIndex        =   26
            Top             =   914
            Width           =   2580
         End
         Begin VB.Label lbl��ֵ 
            BackStyle       =   0  'Transparent
            Caption         =   "��ֵ:"
            Height          =   195
            Left            =   225
            TabIndex        =   25
            Top             =   1380
            Width           =   2580
         End
         Begin VB.Label lbl��ֵ 
            BackStyle       =   0  'Transparent
            Caption         =   "��ֵ:"
            Height          =   195
            Left            =   225
            TabIndex        =   24
            Top             =   1146
            Width           =   2580
         End
         Begin VB.Label lbl�������� 
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   23
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   22
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox pic������Ϣ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   135
         ScaleHeight     =   1950
         ScaleWidth      =   2790
         TabIndex        =   13
         Top             =   135
         Width           =   2790
         Begin VB.PictureBox imgBRXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2460
            Picture         =   "frm���ֽ���༭.frx":00AA
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   37
            Top             =   80
            Width           =   255
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lblסԺ�� 
            BackStyle       =   0  'Transparent
            Caption         =   "ס Ժ ��:"
            Height          =   195
            Left            =   225
            TabIndex        =   19
            Top             =   684
            Width           =   2580
         End
         Begin VB.Label lblסԺ���� 
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����:"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   918
            Width           =   2580
         End
         Begin VB.Label lbl��Ժ���� 
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ����:"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   1152
            Width           =   2580
         End
         Begin VB.Label lbl���� 
            BackStyle       =   0  'Transparent
            Caption         =   "��   ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   16
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label lblסԺҽʦ 
            BackStyle       =   0  'Transparent
            Caption         =   "סԺҽʦ:"
            Height          =   195
            Left            =   225
            TabIndex        =   15
            Top             =   1386
            Width           =   2580
         End
         Begin VB.Label lbl��Ŀ���� 
            BackStyle       =   0  'Transparent
            Caption         =   "��Ŀ����:"
            Height          =   195
            Left            =   225
            TabIndex        =   14
            Top             =   1620
            Width           =   2580
         End
      End
   End
   Begin VB.PictureBox picRight 
      Height          =   7080
      Left            =   3145
      Picture         =   "frm���ֽ���༭.frx":00FF
      ScaleHeight     =   7020
      ScaleWidth      =   8280
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   8340
      Begin VB.ComboBox ComProName 
         Height          =   300
         Left            =   2805
         TabIndex        =   49
         Top             =   825
         Width           =   2070
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   825
         Width           =   1320
      End
      Begin VB.TextBox txt��ע 
         Height          =   300
         Left            =   5460
         TabIndex        =   7
         Tag             =   "��ע"
         Top             =   825
         Width           =   2550
      End
      Begin VB.CheckBox chk�����޸� 
         Caption         =   "�����޸�(&R)"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6705
         TabIndex        =   5
         Top             =   443
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid fgMain 
         Height          =   5820
         Left            =   -45
         TabIndex        =   8
         Top             =   1200
         Width           =   8310
         _cx             =   14658
         _cy             =   10266
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16763080
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   14737632
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   4
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm���ֽ���༭.frx":0318
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin zl9CISAudit.tipPopup tipPopup1 
            Height          =   540
            Left            =   1935
            Top             =   4800
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   953
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txt���� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFE0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3915
            TabIndex        =   36
            Text            =   "333"
            Top             =   945
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ListBox lst���� 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   3960
            TabIndex        =   35
            Top             =   1485
            Visible         =   0   'False
            Width           =   1905
         End
      End
      Begin VB.TextBox txtNo 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   750
         TabIndex        =   1
         Top             =   435
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   2295
         TabIndex        =   50
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   15
         TabIndex        =   47
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lbl��ע 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         Height          =   180
         Left            =   5040
         TabIndex        =   6
         Top             =   885
         Width           =   360
      End
      Begin VB.Label labNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&NO."
         Height          =   180
         Left            =   435
         TabIndex        =   0
         Top             =   495
         Width           =   270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���ֽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   90
         Width           =   1095
      End
      Begin VB.Label lbl�÷� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4230
         TabIndex        =   32
         Top             =   435
         Width           =   600
      End
      Begin VB.Label lbl������ 
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   2970
         TabIndex        =   31
         Top             =   502
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������:"
         Height          =   180
         Left            =   2280
         TabIndex        =   2
         Top             =   495
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�÷�:"
         Height          =   180
         Left            =   3765
         TabIndex        =   3
         Top             =   495
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ȼ�:"
         Height          =   180
         Left            =   5025
         TabIndex        =   4
         Top             =   495
         Width           =   450
      End
      Begin VB.Label lbl�ȼ� 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   5490
         TabIndex        =   33
         Top             =   420
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8505
      TabIndex        =   9
      Top             =   7380
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9750
      TabIndex        =   10
      Top             =   7380
      Width           =   1100
   End
   Begin VB.Label LabStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "12.25%"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4545
      TabIndex        =   46
      Top             =   7440
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label labBar 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2745
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   2670
      X2              =   7185
      Y1              =   7740
      Y2              =   7740
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   2670
      X2              =   7215
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   15
      X2              =   11520
      Y1              =   7335
      Y2              =   7335
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   0
      X2              =   11520
      Y1              =   7230
      Y2              =   7230
   End
   Begin VB.Image imgOpen_White 
      Height          =   225
      Left            =   780
      Picture         =   "frm���ֽ���༭.frx":0497
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgClose_White 
      Height          =   225
      Left            =   1185
      Picture         =   "frm���ֽ���༭.frx":04F9
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   105
      Picture         =   "frm���ֽ���༭.frx":054E
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   465
      Picture         =   "frm���ֽ���༭.frx":05A3
      Top             =   7980
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   0
      Picture         =   "frm���ֽ���༭.frx":05F2
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   0
      Picture         =   "frm���ֽ���༭.frx":07B2
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label pbrBar 
      Height          =   240
      Left            =   2670
      TabIndex        =   45
      Top             =   8625
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "frm���ֽ���༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private mfrmArchiveView As frmArchiveView
Private m_lng���ID         As Long
Private m_lng����ID         As Long
Private m_lng��ҳID         As Long
Private m_lng����ID         As Long
Private m_lng����ID         As Long
Private m_str��ʽ           As String     '��ӡ��޸ġ�����
Private m_blnModed          As Boolean
Private edRow%, edCol%, edKey%
Private m_lngOldSJID        As Long         '�ɵ��ϼ�ID
Private m_lngCurSJID        As Long         '�ϼ�ID
Private m_bln��α༭       As Boolean
Private zlCheck             As New clsCheck
Private mblnStop            As Boolean
Dim mbln��Ŀ������          As Boolean
Public Event AferSaveData()

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

'==============================================================================
'=���ܣ� ������ʾ
'==============================================================================
Public Sub ShowForm(��ʽ As String, ���ID As Long, ����ID As Long, ��ҳID As Long, ����ID As Long, ����ID As Long)
    Dim rsTemp      As ADODB.Recordset
    
    On Error GoTo errH
    
    m_bln��α༭ = False
    
    m_blnModed = False
    m_str��ʽ = ��ʽ    '���/�޸�/����
    m_lng���ID = ���ID
    m_lng����ID = ����ID
    m_lng��ҳID = ��ҳID
    m_lng����ID = ����ID
    
    '��ʼ������
    Select Case ��ʽ
        Case "����"
            Me.Caption = "��������"
        Case "�޸�"
            Me.Caption = "�޸����ֽ��"
        Case "����"
            Me.Caption = "��������"
    End Select
    
    If m_str��ʽ = "����" Or m_str��ʽ = "����" Then
        '��m_lng����ID��ΪĬ�Ϸ���ID
        gstrSQL = "select ID from �������ַ��� where ����= [1] and ѡ�� = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "סԺ", 1)
        
        If rsTemp.EOF Then
            MsgBox "���ڡ����ֱ�׼ά����������Ĭ��ѡ�õ����ַ�����", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        Else
            m_lng����ID = rsTemp.Fields(0).Value
        End If
    Else
        m_lng����ID = ����ID
    End If
    
    Call Fill��������
    Call Fill���ֱ�׼
    Call Fill������Ŀ
    
    If m_str��ʽ = "�޸�" Then
        Call Fill���ֽ��
    End If
    Me.Tag = "��ɳ�ʼ��"
    fgMain_CellChanged 0, 0
    If Me.Visible = False Then Me.Show
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������ַ���ID���������ֱ�׼����
'=       �����˽����ı������ݵĳ�ʼ��  m_lng����ID   m_lng����ID  m_lng��ҳID
'=
'=       m_lng����ID  m_lng��ҳID  =>������Ϣ
'==============================================================================
Private Sub Fill���ֱ�׼()

    Dim rsTemp          As ADODB.Recordset
    Dim lngIndex        As Long
    Dim i               As Long

    On Error GoTo errH
        
    lst����.Clear
    lst����.AddItem "1 - ��"
    lst����.AddItem "2 - ����"
    txt����.Text = ""
    
    gstrSQL = "select ����,סԺ��,��Ժ����,סԺҽʦ,��Ŀ���� from ��������������ͼ where ����ID=[1] And ��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID, m_lng��ҳID)
    
    If rsTemp.EOF Then
        MsgBox "û��ѡ����", vbInformation, gstrSysName
        CmdCancel_Click
        Exit Sub
    Else
        lbl���� = "��   ��:" & NVL(rsTemp("����"))
        txtNo = NVL(rsTemp("����"))
        lblסԺ�� = "ס Ժ ��:" & NVL(rsTemp("סԺ��"))
        lblסԺ���� = "סԺ����:" & m_lng��ҳID
        lbl��Ժ���� = "��Ժ����:" & NVL(rsTemp("��Ժ����"))
        lblסԺҽʦ = "סԺҽʦ:" & NVL(rsTemp("סԺҽʦ"))
        lbl��Ŀ���� = "��Ŀ����:" & NVL(rsTemp("��Ŀ����"))
    End If
    rsTemp.Close
    
    'm_lng����ID:������Ϣ
    gstrSQL = "select ����,����,�ܷ�,��ֵ,��ֵ from �������ַ��� where ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
    If rsTemp.EOF Then
        MsgBox "δ֪�����ַ�����", vbInformation, gstrSysName
        CmdCancel_Click
        Exit Sub
    Else
        lbl�������� = "����:" & NVL(rsTemp("����"))
        lbl���� = "����:" & NVL(rsTemp("����"))
        lbl�ܷ� = "�ܷ�:" & NVL(rsTemp("�ܷ�"))
        lbl��ֵ = "��ֵ:" & NVL(rsTemp("��ֵ"))
        lbl��ֵ = "��ֵ:" & NVL(rsTemp("��ֵ"))
    End If
    rsTemp.Close
    
    '������Ϣ
    If m_lng���ID <> 0 And m_str��ʽ = "�޸�" Then
        gstrSQL = "" & _
            "   Select Id, ����id, ��ҳid, ����id, �ܷ�, �ȼ�, �����޸�,��������, ������, ����ʱ��, �����, ���ʱ��, ��ע " & _
            "   From �������ֽ�� " & _
            "   Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng���ID)
        
        If Not rsTemp.EOF Then
            lbl������ = NVL(rsTemp("������"), "")
            If InStr(gstrPrivs, "�޸���������") = 0 And UCase(lbl������) <> UCase(gstrUserName) Then  '�����п��ҹ���
                cmdOK.Enabled = False
            End If
            lbl�÷� = NVL(rsTemp("�ܷ�"), 0)
            lbl�ȼ� = NVL(rsTemp("�ȼ�"), "")
            chk�����޸�.Value = IIf(NVL(rsTemp("�����޸�"), 0) = 0, vbUnchecked, vbChecked)
            txt��ע.Text = NVL(rsTemp!��ע)
            If NVL(rsTemp!��������) = "" Then
                cbo.ListIndex = 0
            Else
                On Error Resume Next
                cbo.Text = NVL(rsTemp!��������)
            End If
        End If
        rsTemp.Close
    Else
        lbl������ = gstrUserName
        lbl�÷� = ""
        lbl�ȼ� = ""
    End If
    
    If m_bln��α༭ = True Then    '�ڶ��������Ͳ�����д���ֱ�׼�ˣ�Ĭ�ϱ�׼��
        If fgMain.Rows > 1 Then
            fgMain.Row = 1
            fgMain.ShowCell 1, 4
            fgMain.SetFocus
        End If
        Exit Sub
    End If
    
    'ȷ������
    Dim bln�۷��� As Boolean, intSign As Long
    gstrSQL = "select ���� from �������ַ��� where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
    
    bln�۷��� = True
    If Not rsTemp.EOF Then
        bln�۷��� = IIf(NVL(rsTemp("����"), "�ӷ���") = "�ӷ���", False, True)
    End If
    rsTemp.Close
    
    If bln�۷��� Then
        intSign = -1
    Else
        intSign = 1
    End If

    With fgMain
        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        '��������
        .Cols = 11
        .Cell(flexcpText, 0, 0) = "��Ŀ"
        .Cell(flexcpText, 0, 1) = "��׼��ֵ"
        .Cell(flexcpText, 0, 2) = "ȱ������"
        .Cell(flexcpText, 0, 3) = "���ֱ�׼"
        .Cell(flexcpText, 0, 4) = "����"
        .Cell(flexcpText, 0, 5) = "�ɷ��޸�"
        .Cell(flexcpText, 0, 6) = "ID"
        .Cell(flexcpText, 0, 7) = "�ϼ�ID"
        .Cell(flexcpText, 0, 8) = "����ID"
        .Cell(flexcpText, 0, 9) = "��ע"
        .Cell(flexcpText, 0, 10) = "����ȼ�"
        .ExtendLastCol = True
        'ȷ����������
        If m_lng����ID < 1 Then .Redraw = flexRDDirect: Exit Sub
        
        gstrSQL = "" & _
            "   Select �ϼ����, ���, Id, �ϼ�id, ����id, ��Ŀ, ��׼��ֵ, ����Ҫ��, ȱ������, �۷ֱ�׼, ����,����ȼ� " & _
            "   From �������ֱ�׼��ͼ " & _
            "   Where ����='��' and ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
        .FocusRect = flexFocusSolid
        .Rows = rsTemp.RecordCount + 1
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, 0) = IIf(IsNull(rsTemp.Fields("��Ŀ")), "", rsTemp.Fields("��Ŀ"))
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("��׼��ֵ")), " ", Format(rsTemp.Fields("��׼��ֵ"), "####��"))
            .Cell(flexcpText, i, 2) = IIf(IsNull(rsTemp.Fields("ȱ������")), "", rsTemp.Fields("ȱ������"))
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("�۷ֱ�׼")), "", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�׼�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�Ҽ�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "����", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "������", rsTemp.Fields("�۷ֱ�׼"))))))
            .Cell(flexcpText, i, 4) = ""
            If intSign = 1 Then
                .Cell(flexcpForeColor, i, 4) = RGB(0, 0, 255)
            Else
                .Cell(flexcpForeColor, i, 4) = RGB(255, 0, 0)
            End If
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            .Cell(flexcpText, i, 5) = ""
            .Cell(flexcpForeColor, i, 5) = RGB(0, 0, 0)
            .Cell(flexcpText, i, 6) = IIf(IsNull(rsTemp.Fields("ID")), "", rsTemp.Fields("ID"))
            .Cell(flexcpText, i, 7) = IIf(IsNull(rsTemp.Fields("�ϼ�ID")), "", rsTemp.Fields("�ϼ�ID"))
            .Cell(flexcpText, i, 8) = IIf(IsNull(rsTemp.Fields("����ID")), "", rsTemp.Fields("����ID"))
            .Cell(flexcpText, i, 9) = ""
            .Cell(flexcpText, i, 10) = NVL(rsTemp.Fields!����ȼ�)
            rsTemp.MoveNext
            i = i + 1
        Loop
        '�Զ�����
        .WordWrap = True
        '�ϲ���Ԫ��
        .MergeCells = 2
        .MergeCol(.ColIndex("��Ŀ")) = True
        .MergeCol(.ColIndex("��׼��ֵ")) = True
        '��������
        .ColAlignment(.ColIndex("��Ŀ")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("��׼��ֵ")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("���ֱ�׼")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("�ɷ��޸�")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("��ע")) = flexAlignLeftCenter
        
        '���ص�Ԫ��
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("�ϼ�ID")) = 0
        .ColWidth(.ColIndex("����ID")) = 0
        '�������
        .ColWidth(.ColIndex("��Ŀ")) = 600
        .ColWidth(.ColIndex("��׼��ֵ")) = 600
        .ColWidth(.ColIndex("ȱ������")) = 3200
        .ColWidth(.ColIndex("���ֱ�׼")) = 950
        .ColWidth(.ColIndex("����")) = 1000
        .ColWidth(.ColIndex("�ɷ��޸�")) = 850
        '�и�����
        .RowHeightMin = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("ȱ������")
        .SelectionMode = flexSelectionFree
        .AllowBigSelection = False
        
        .Editable = flexEDKbdMouse   '�ɱ༭

        .Redraw = flexRDBuffered
        'ѡ����ǰ����
    End With
    If fgMain.Rows > 1 Then
        fgMain.Row = 1
        fgMain.ShowCell 1, 4
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

'==============================================================================
'=���ܣ� װ���Ӧ��ҳ�����ֽ��
'==============================================================================
Private Function Fill���ֽ��() As Boolean
    Dim rs              As ADODB.Recordset
    Dim lngJGID         As Long
    Dim i               As Long
    
    On Error GoTo errH
    
    fgMain.Redraw = flexRDNone
    
    For i = 1 To fgMain.Rows - 1
        fgMain.Cell(flexcpText, i, 4) = ""
        fgMain.Cell(flexcpText, i, 5) = ""
    Next
    
    'ȷ������
    Dim bln�۷��� As Boolean, intSign As Long
    gstrSQL = "select ���� from �������ַ��� where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
    
    bln�۷��� = True
    If Not rs.EOF Then
        bln�۷��� = IIf(NVL(rs("����"), "�ӷ���") = "�ӷ���", False, True)
    End If
    rs.Close
    
    If bln�۷��� Then
        intSign = -1
    Else
        intSign = 1
    End If
    
    gstrSQL = "" & _
        "   select  A.ID,A.��Ŀ,A.��׼��ֵ,A.����Ҫ��,A.ȱ������,A.�۷ֱ�׼," & _
        "           (select decode(ȱ�ݵȼ�,null,to_CHAR(�������),ȱ�ݵȼ�) from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[1]) as ����," & _
        "           (select �ɷ��޸� from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[1]) as �ɷ��޸�," & _
        "           (select ��ע from ����������ϸ where ���ֱ�׼ID=A.ID and ����ID=[1]) as ��ע" & _
        "   From �������ֱ�׼��ͼ A " & _
        "   Where A.����='��' and A.����ID=(select B.����ID from �������ֽ�� B where B.ID=[1]) " & _
        "   Order by A.�ϼ�ID,A.ID "
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng���ID)
        
    If Not rs.EOF Then
        For i = 1 To fgMain.Rows - 1
            rs.MoveFirst
            rs.Find "ID=" & Val(fgMain.Cell(flexcpText, i, 6))
            If Not rs.EOF Then
                If Not IsNull(rs("����")) Then
                    Select Case rs("����")
                    Case "��", "��", "��"
                        fgMain.Cell(flexcpText, i, 4) = rs("����").Value + "��"
                    Case "��"
                        fgMain.Cell(flexcpText, i, 4) = "������"
                    Case Else
                        fgMain.Cell(flexcpText, i, 4) = IIf(Abs(NVL(rs("����").Value, 0)) < 1, Format(Abs(NVL(rs("����").Value, 0)), "0.0"), Abs(NVL(rs("����").Value, 0)))
                        If intSign = -1 Then
                            fgMain.Cell(flexcpForeColor, i, 4) = RGB(255, 0, 0)
                        Else
                            fgMain.Cell(flexcpForeColor, i, 4) = RGB(0, 0, 255)
                        End If
                    End Select
                End If
                If Not IsNull(rs("�ɷ��޸�")) Then
                    If rs("�ɷ��޸�") = 1 Then
                        fgMain.Cell(flexcpText, i, 5) = "��"
                    End If
                End If
                fgMain.Cell(flexcpText, i, 9) = NVL(rs!��ע)
            End If
        Next
    End If
    
    fgMain.Redraw = flexRDBuffered
    If fgMain.Rows > 1 Then
        fgMain.Row = 1
        fgMain.ShowCell 1, 4
    End If

    Fill���ֽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Fill���ֽ�� = False
End Function

Private Sub Fill��������()
    Dim rs As New ADODB.Recordset
    On Error GoTo errH
    gstrSQL = "" & _
        "Select ����,����,����,ȱʡ��־ From ��������"
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With cbo
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = 1
        
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .AddItem zlCommFun.NVL(rs!����)
                 .ItemData(.NewIndex) = .NewIndex + 1

                rs.MoveNext
            Loop
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Fill������Ŀ()
    Dim rs As New ADODB.Recordset
    On Error GoTo errH
    gstrSQL = "" & _
        "    Select A.ID,A.���� From �������ֱ�׼ A,�������ַ��� B Where A.����ID= B.ID And B.ѡ��=1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With ComProName
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = 1
        
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .AddItem zlCommFun.NVL(rs!����)
                 .ItemData(.NewIndex) = .NewIndex + 1

                rs.MoveNext
            Loop
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
        
    End With
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume Next
    End If
End Sub

'==============================================================================
'=���ܣ� �����޸�ֵ��������
'==============================================================================
Private Sub chk�����޸�_Click()
    On Error GoTo errH
    If chk�����޸�.Value = vbChecked Then
        chk�����޸�.FontBold = True
    Else
        chk�����޸�.FontBold = False
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Զ��������Ӧ�ķ�ֵ��д�����ֱ�׼
'==============================================================================
Private Sub cmdAuto_Click()
Dim lngLoop As Long, strID As String, strSQL As String
Dim strReturn As String, strMid As String, strAlidin As String
    Dim rsTemp      As ADODB.Recordset
    
    On Error GoTo errH
    
    If fgMain.Rows = 1 Then
        zlCheck.Msg_OK "�����ַ������������֣�", vbCritical
        Exit Sub
    End If
    If zlCheck.Msg_OKC("ȷ�Ͻ����Զ����ּ�����") Then Exit Sub
    
    cmdAuto.Visible = False
    cmdStop.Visible = True
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    labBar.Width = 0
    labBar.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    LabStatus.Visible = True
    
    DoEvents
    '��ȡ���ڵ�ǰ�еļ�¼����
    For lngLoop = 1 To fgMain.Rows - 1
        LabStatus.Caption = Format(Round((lngLoop / (fgMain.Rows - 1)) * 100, 2), "0.00") & " %"
        labBar.Width = lngLoop * pbrBar.Width / (fgMain.Rows - 1)
        If mblnStop Then
            cmdAuto.Visible = True
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            
            cmdStop.Visible = False
            labBar.Visible = False
            Line3.Visible = False
            Line4.Visible = False
            LabStatus.Visible = False
            mblnStop = False
            Call zlCheck.Msg_OK("�����Զ�������;ȡ��������ɲ������֣�", vbCritical)
            Exit Sub
        End If
        strID = fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ID"))
        
        strSQL = "select �ж�����,����Դ from �������ֱ�׼ where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID)
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            strSQL = "" & rsTemp.Fields!�ж�����
            If strSQL <> "" Then
                If rsTemp!����Դ = 0 Then
                    strSQL = CheckAuditSql_OUT(strSQL, m_lng����ID, m_lng��ҳID)
                    Set rsTemp = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSQL, "'", "''") & "') from dual", Me.Caption)
                ElseIf gobjEmr Is Nothing Then
                    MsgBox "����δ��װ������������ܽ������֣����飡", vbInformation, gstrSysName
                    mblnStop = True
                ElseIf Not gobjEmr Is Nothing Then
                    If strMid = "" Then Call GetEMR_MID_ALIDIN(m_lng����ID, m_lng��ҳID, strMid, strAlidin) 'ȡ�²�������ID,�ID
                    strSQL = Replace(rsTemp!�ж�����, "[MID]", ":mid")
                    strSQL = Replace(rsTemp!�ж�����, "[ALIDIN]", ":alidin")
                    strReturn = gobjEmr.OpenSQLRecordset(strSQL, IIf(strMid = "", "", strMid & "^" & DbType.T_String & "^mid") & IIf(strAlidin = "", "", IIf(strMid = "", "", "|") & strAlidin & "^" & DbType.T_String & "^alidin"), rsTemp)
                    If strReturn <> "" Then Set rsTemp = New ADODB.Recordset
                End If
                
                If Not zlCheck.Connection_ChkRsState(rsTemp) Then
                    If InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]") = 0 Then
                        fgMain.TextMatrix(lngLoop, fgMain.ColIndex("����")) = "" & rsTemp.Fields(0)
                    Else
                        fgMain.TextMatrix(lngLoop, fgMain.ColIndex("����")) = 0
                    End If
                End If
            End If
        End If
        DoEvents
    Next
    zlCheck.Msg_OK ("�����Զ����ֳɹ���")
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    labBar.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    LabStatus.Visible = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call zlCheck.Msg_OK("�����Զ�����ʧ�ܣ�", vbCritical)
    cmdAuto.Visible = True
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdStop.Visible = False
    labBar.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    LabStatus.Visible = False
End Sub

'==============================================================================
'=���ܣ� ֹͣ�Զ�����
'==============================================================================
Private Sub cmdStop_Click()
    On Error GoTo errH
    
    mblnStop = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���Ĳ���
'==============================================================================
Private Sub cmdViewArchive_Click()
    On Error GoTo errH
    If mfrmArchiveView Is Nothing Then Set mfrmArchiveView = New frmArchiveView
    Call mfrmArchiveView.ShowArchive(Me, m_lng����ID, m_lng��ҳID, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ComProName_Click()
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo errH
    
    lngRow = 0
    If ComProName.Locked Then Exit Sub

    If fgMain.ColIndex("ȱ������") = -1 Then Exit Sub
  
    '��ȡ���ڵ�ǰ�еļ�¼����
    For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
        If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ȱ������"))), UCase(ComProName.Text)) > 0 Then
            lngRow = lngLoop
            Exit For
        End If
    Next
    '��ȡС�ڵ�ǰ�еļ�¼����
    If lngRow = 0 Then
        For lngLoop = 0 To fgMain.Row
            If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ȱ������"))), UCase(ComProName.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
    End If
    If fgMain.Rows > 1 And lngRow >= 1 Then fgMain.Row = lngRow
'        fgMain.Cell lngRow
    fgMain.ShowCell lngRow, 4
    Call LocationObj(ComProName)
 
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ComProName_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    Dim strTmpProName As String
    
    On Error GoTo errH
    
    lngRow = 0
    If ComProName.Locked Then Exit Sub

    If fgMain.ColIndex("ȱ������") = -1 Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        
        If zlCommFun.IsNumOrChar(ComProName.Text) Then
            For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
                If InStr(UCase(zlCommFun.SpellCode(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ȱ������")))), UCase(ComProName.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
            '��ȡС�ڵ�ǰ�еļ�¼����
            If lngRow = 0 Then
                For lngLoop = 0 To fgMain.Row
                    If InStr(UCase(zlCommFun.SpellCode(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ȱ������")))), UCase(ComProName.Text)) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
        Else
            For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
                If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ȱ������"))), UCase(ComProName.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
            '��ȡС�ڵ�ǰ�еļ�¼����
            If lngRow = 0 Then
                For lngLoop = 0 To fgMain.Row
                    If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex("ȱ������"))), UCase(ComProName.Text)) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then fgMain.Row = lngRow
'        fgMain.Cell lngRow
        fgMain.ShowCell lngRow, 4
        Call LocationObj(ComProName)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������޸ı༭
'==============================================================================
Private Sub fgMain_Click()
    On Error GoTo errH
    
    If fgMain.Col = 5 Then
        If fgMain.Cell(flexcpText, fgMain.Row, 5) = "" Then
            fgMain.Cell(flexcpText, fgMain.Row, 5) = "��"
        Else
            fgMain.Cell(flexcpText, fgMain.Row, 5) = ""
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����༭�в������롰'��
'==============================================================================
Private Sub fgMain_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error GoTo errH
    
    If KeyAscii = Asc("'") Then
       KeyAscii = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���絥��ʱ��m_lngCurSJID��Ŀ������ֵ
'==============================================================================
Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 6)))      '��ȡID
End Sub

'==============================================================================
'=���ܣ� �������б䶯ʱ��ȡ
'==============================================================================
Private Sub fgMain_RowColChange()
    Dim lngID               As Long
    Dim lngCurID            As Long
    Dim lngCurSJID          As Long
    
    On Error GoTo errH
    
    If fgMain.Row < 0 Then
        lngCurSJID = 0
        lngCurID = 0
        Exit Sub
    End If
    
    lngCurID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 6)))         '��ȡID
    lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 7)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 7)))       '��ȡ�ϼ�ID
    
    If lngCurSJID = 0 Then
        lngID = lngCurID
    Else
        lngID = lngCurSJID
    End If
    
    Show����Ҫ�� lngID, fgMain.Cell(flexcpText, fgMain.Row, 0), fgMain.Cell(flexcpText, fgMain.Row, 1)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������º����ַ�����
'==============================================================================
Private Sub fgMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    On Error GoTo errH
    
    If Col = 9 Then
        If zlCommFun.ActualLen(Trim(fgMain.EditText)) > 50 Then
            MsgBox "������ı�ע������25�����ֻ�50���ַ�,���ܼ���!"
            Cancel = True
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڵ�����ʼ��
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    Call InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ҳ���ʼ��
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    '��ȡϵͳ�������Ƿ��Ŀ���������
    mbln��Ŀ������ = Val(zlDatabase.GetPara(91, glngSys, 0)) = 1
    m_lngOldSJID = -1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڱ仯
'==============================================================================
Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 8175 Then
        Me.Height = 8175
    End If
    If Me.Width < 11520 Then
        Me.Width = 11640
    End If
    With txt��ע '
        .Width = ScaleWidth - .Left - 50
    End With
    With cmdCancel
        .Top = ScaleHeight - .Height - 85
        .Left = ScaleWidth - .Width - 100
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - .Width - 50
    End With
    With cmdAuto
        .Top = cmdOK.Top
        .Left = cmdOK.Left - .Width - 50
    End With
    With cmdStop
        .Top = cmdOK.Top
        .Left = cmdOK.Left - .Width - 50
    End With
    With cmdHelp
        .Top = cmdCancel.Top
    End With
    With cmdViewArchive
        .Top = cmdCancel.Top
    End With
    With pbrBar
        .Top = cmdHelp.Top + 30
        .Left = cmdViewArchive.Left + cmdViewArchive.Width + 200
        .Width = cmdStop.Left - cmdViewArchive.Left - cmdViewArchive.Width - 400
    End With
    With Line1
        .Y1 = cmdCancel.Top - 85
        .y2 = .Y1
        .X1 = 0
        .x2 = ScaleWidth
    End With
    With Line2
        .Y1 = Line1.Y1 + 30
        .y2 = .Y1
        .X1 = 0
        .x2 = ScaleWidth
    End With
    
    With picLeft
        .Height = Line1.Y1 - 85 - .Top
    End With
    
    With fgMain
        .Width = ScaleWidth - .Left - 50
        .Height = Line1.Y1 - .Top - 85
    End With

    With chk�����޸�
        .Left = ScaleWidth - .Width
    End With
    With picRight
        .Width = ScaleWidth - .Left
        .Height = Line1.Y1 - .Top - 85
    End With
    With Line3
        .Y1 = pbrBar.Top - 10
        .y2 = pbrBar.Top - 10
        .X1 = pbrBar.Left
        .x2 = pbrBar.Left + pbrBar.Width
    End With
    With Line4
        .Y1 = pbrBar.Top + pbrBar.Height + 40
        .y2 = pbrBar.Top + pbrBar.Height + 40
        .X1 = pbrBar.Left
        .x2 = pbrBar.Left + pbrBar.Width
    End With
    With LabStatus
        .Move pbrBar.Left + pbrBar.Width / 2 - 50, pbrBar.Top + 50
    End With
    With labBar
        .Move pbrBar.Left, pbrBar.Top + 20
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmArchiveView
    Set mfrmArchiveView = Nothing
    Set zlCheck = Nothing
End Sub
'==============================================================================
'=���ܣ� ������Ϣ����
'==============================================================================
Private Sub imgBRXX_Click()

    On Error GoTo errH
    
    If imgBRXX.Tag = "" Then
        imgBRXX.Tag = "Opened"
        imgBRXX.Picture = imgOpen_White.Picture
        pic������Ϣ.Height = 340
    Else
        imgBRXX.Tag = ""
        imgBRXX.Picture = imgClose_White.Picture
        pic������Ϣ.Height = 1950
    End If
    imgBRXX.Refresh
    picLeft_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�������Ϣ�仯
'==============================================================================
Private Sub imgBRXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If X >= 0 And X <= imgBRXX.ScaleWidth And Y >= 0 And Y <= imgBRXX.ScaleHeight Then
        SetCapture imgBRXX.hWnd
        '������룡����
        imgBRXX.Line (0, 0)-(imgBRXX.ScaleWidth - Screen.TwipsPerPixelX, imgBRXX.ScaleHeight - Screen.TwipsPerPixelY), vbWhite, B
    Else
        '����Ƴ�������
        imgBRXX.Cls
        ReleaseCapture
    End If
End Sub

'==============================================================================
'=���ܣ�������Ϣ����
'==============================================================================
Private Sub imgFAXX_Click()
On Error Resume Next
    If imgFAXX.Tag = "" Then
        imgFAXX.Tag = "Opened"
        imgFAXX.Picture = imgOpen.Picture
        pic������Ϣ.Height = 340
    Else
        imgFAXX.Tag = ""
        imgFAXX.Picture = imgClose.Picture
        pic������Ϣ.Height = 1695
    End If
    imgFAXX.Refresh
    picLeft_Resize
End Sub

'==============================================================================
'=���ܣ�������Ϣ�仯
'==============================================================================
Private Sub imgFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If X >= 0 And X <= imgFAXX.ScaleWidth And Y >= 0 And Y <= imgFAXX.ScaleHeight Then
        SetCapture imgFAXX.hWnd
        '������룡����
        imgFAXX.Line (0, 0)-(imgFAXX.ScaleWidth - Screen.TwipsPerPixelX, imgFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '����Ƴ�������
        imgFAXX.Cls
        ReleaseCapture
    End If
End Sub

'==============================================================================
'=���ܣ���Ŀ��Ϣ����
'==============================================================================
Private Sub imgXMXX_Click()
On Error Resume Next
    If imgXMXX.Tag = "" Then
        imgXMXX.Tag = "Opened"
        imgXMXX.Picture = imgOpen.Picture
        pic��Ŀ��Ϣ.Height = 340
    Else
        imgXMXX.Tag = ""
        imgXMXX.Picture = imgClose.Picture
        pic��Ŀ��Ϣ.Height = Abs(picLeft.ScaleHeight - pic������Ϣ.Height - pic������Ϣ.Height - 135 * 4)
    End If
    imgXMXX.Refresh
    picLeft_Resize
End Sub

'==============================================================================
'=���ܣ���Ŀ��Ϣ�仯
'==============================================================================
Private Sub imgXMXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If X >= 0 And X <= imgXMXX.ScaleWidth And Y >= 0 And Y <= imgXMXX.ScaleHeight Then
        SetCapture imgXMXX.hWnd
        '������룡����
        imgXMXX.Line (0, 0)-(imgXMXX.ScaleWidth - Screen.TwipsPerPixelX, imgXMXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '����Ƴ�������
        imgXMXX.Cls
        ReleaseCapture
    End If
End Sub

'==============================================================================
'=���ܣ�����˫���༭
'==============================================================================
Private Sub lst����_DblClick()
    fgMain.SetFocus
    If fgMain.Row = fgMain.Rows - 1 Then
        If cmdOK.Enabled Then cmdOK.SetFocus
        Exit Sub
    End If
    fgMain.Row = fgMain.Row + 1
    fgMain.ShowCell fgMain.Row, 4
End Sub

'==============================================================================
'=���ܣ� ���ֱ༭���
'==============================================================================
Private Sub lst����_LostFocus()
    On Error GoTo errH
    
    If lst����.ListIndex = 0 Then
        fgMain.TextMatrix(edRow, edCol) = ""
    Else
        fgMain.TextMatrix(edRow, edCol) = fgMain.Cell(flexcpText, edRow, 3)
    End If
    lst����.Visible = False
    edKey = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ȡ��
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo errH
    
    If m_bln��α༭ Then
        Moded = True
    Else
        Moded = False
    End If
    Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo errH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ȷ����������
'==============================================================================
Private Sub CmdOK_Click()
    Dim strT            As String
    Dim r               As Long
    Dim lngID           As Long
    Dim lng��ϸID       As Long
    Dim lng�ɷ��޸�     As Long
    
    On Error GoTo errH
    If zlCommFun.ActualLen(Trim(txt��ע.Text)) > 50 Then
        MsgBox "������ı�ע������25�����ֻ�50���ַ�,���ܼ���!"
        txt��ע.SelStart = 1
        txt��ע.SelLength = 100
        If txt��ע.Enabled Then txt��ע.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    '������
    If m_str��ʽ = "����" Or m_str��ʽ = "�޸�" Then  '�����ǰ�����ֽ��
        gstrSQL = "ZL_�������ֽ��_Delete(" & m_lng���ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    lngID = zlDatabase.GetNextId("�������ֽ��")
    'Zl_�������ֽ��_Insert
    gstrSQL = "ZL_�������ֽ��_Insert("
    '  Id_In       In �������ֽ��.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  ����id_In   In �������ֽ��.����id%Type,
    gstrSQL = gstrSQL & "" & m_lng����ID & ","
    '  ��ҳid_In   In �������ֽ��.��ҳid%Type,
    gstrSQL = gstrSQL & "" & m_lng��ҳID & ","
    '  ����id_In   In �������ֽ��.����id%Type,
    gstrSQL = gstrSQL & "" & m_lng����ID & ","
    '  �ܷ�_In     In �������ֽ��.�ܷ�%Type,
    gstrSQL = gstrSQL & "" & Val(lbl�÷�) & ","
    '  �ȼ�_In     In �������ֽ��.�ȼ�%Type,
    gstrSQL = gstrSQL & "'" & IIf(lbl�ȼ� = "���ϸ�", "��", lbl�ȼ�) & "',"
    '  ��������_In In �������ֽ��.��������%Type,
    gstrSQL = gstrSQL & "" & IIf(Trim(cbo.Text) = "", "NULL", "'" & Trim(cbo.Text) & "'") & ","
    '  ��ע_In     In �������ֽ��.��ע%Type,
    gstrSQL = gstrSQL & "" & IIf(Trim(txt��ע.Text) = "", "NULL", "'" & Trim(txt��ע.Text) & "'") & ","
    '  ������_In   In �������ֽ��.������%Type,
    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
    '  ����ʱ��_In In �������ֽ��.����ʱ��%Type,
    gstrSQL = gstrSQL & "Sysdate,"
    '  �����_In   In �������ֽ��.�����%Type,
    gstrSQL = gstrSQL & "NULL,"
    '  ���ʱ��_In In �������ֽ��.���ʱ��%Type,
    gstrSQL = gstrSQL & "NULL,"
    '  �����޸�_In In �������ֽ��.�����޸�%Type
    gstrSQL = gstrSQL & "" & IIf(chk�����޸�.Value = vbChecked, "1", "Null") & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
     
    strT = "ZL_����������ϸ_Insert"
    
    For r = 1 To fgMain.Rows - 1
        If Trim(fgMain.Cell(flexcpText, r, 4)) <> "" Or fgMain.Cell(flexcpText, r, 5) <> "" Then
            If fgMain.Cell(flexcpText, r, 5) = "��" Then
                lng�ɷ��޸� = 1
            Else
                lng�ɷ��޸� = 0
            End If
            lng��ϸID = zlDatabase.GetNextId("����������ϸ")    '��������

            gstrSQL = "ZL_����������ϸ_Insert("
            gstrSQL = gstrSQL & "" & lng��ϸID & ","
            gstrSQL = gstrSQL & "" & lngID & ","
            gstrSQL = gstrSQL & "" & fgMain.Cell(flexcpText, r, 6) & ","
            
            Select Case fgMain.Cell(flexcpText, r, 4)
                Case "�׼�"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'��',"
                Case "�Ҽ�"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'��',"
                Case "����"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'��',"
                Case "������"
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "'��',"
                Case "" 'ֻ�С��ɷ��޸ġ�����д�ˣ�
                    gstrSQL = gstrSQL & "null,"
                    gstrSQL = gstrSQL & "null,"
                Case Else
                    gstrSQL = gstrSQL & "" & Abs(Val(fgMain.Cell(flexcpText, r, 4))) & ","
                    gstrSQL = gstrSQL & "null,"
            End Select
            gstrSQL = gstrSQL & "" & lng�ɷ��޸� & ","
            gstrSQL = gstrSQL & "" & IIf(Trim(fgMain.Cell(flexcpText, r, 9)) = "", "Null", "'" & Trim(fgMain.Cell(flexcpText, r, 9)) & "'") & ")"

            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans
    Moded = True
    MsgBox "���ֽ������ɹ���", vbOKOnly + vbInformation, gstrSysName
    RaiseEvent AferSaveData
    If m_str��ʽ = "����" Then
        Call ClearResults
        cmdOK.Enabled = False
        zlControl.TxtSelAll txtNo
        txtNo.SetFocus
        m_bln��α༭ = True
    Else
        Unload Me
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����༭���
'==============================================================================
Private Sub fgMain_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    On Error GoTo errH
    
    If txt����.Visible Then
        txt����.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX, fgMain.CellHeight - Screen.TwipsPerPixelY
    End If
    If lst����.Visible Then
        lst����.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �����ƶ��仯
'==============================================================================
Private Sub fgMain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errH
    
    If txt����.Visible Then
        txt����.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX, fgMain.CellHeight - Screen.TwipsPerPixelY
    End If
    If lst����.Visible Then
        lst����.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��갴��֮ǰ����ֵ��
'==============================================================================
Private Sub fgMain_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    On Error GoTo errH
    edKey = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����༭��ɺ󣬶�̬�ı�÷ֺ͵ȼ�
'==============================================================================
Private Sub fgMain_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errH
    
    If Me.Tag = "" Then Exit Sub
    If fgMain.Rows > 1 Then
        lbl�÷� = Get����
    Else
        lbl�÷� = ""
    End If
    
    If fgMain.Rows > 1 Then
        lbl�ȼ� = Get�ȼ�
    Else
        lbl�ȼ� = ""
    End If
    If lbl�ȼ� = "���ϸ�" Then
        lbl�÷�.Visible = False
    Else
        lbl�÷�.Visible = True
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȡ������ͳ�Ƶķ���
'==============================================================================
Private Function Get����() As Single
    Dim r               As Long
    Dim SUM��Ŀ         As Single
    Dim SUM�ܷ�         As Single
    Dim ��Ŀ            As String
    Dim Num             As Single
    
    On Error GoTo errH
    
    If fgMain.Rows > 1 Then ��Ŀ = fgMain.Cell(flexcpText, 1, 0)
    For r = 1 To fgMain.Rows - 1
        If ��Ŀ <> fgMain.Cell(flexcpText, r, 0) Then
            If SUM��Ŀ > Val(fgMain.Cell(flexcpText, r - 1, 1)) Then
                SUM��Ŀ = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1)))
            ElseIf Abs(SUM��Ŀ) < 0.001 Then
                SUM��Ŀ = 0#
            Else
                SUM��Ŀ = Abs(SUM��Ŀ)
            End If
            If Right(lbl����, 3) = "�۷���" Then
                SUM��Ŀ = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1))) - SUM��Ŀ
            End If
            SUM�ܷ� = SUM�ܷ� + SUM��Ŀ
            SUM��Ŀ = 0#
        End If
        
        Num = Abs(Val(fgMain.Cell(flexcpText, r, 4)))
        SUM��Ŀ = SUM��Ŀ + CDbl(Num)
        ��Ŀ = fgMain.Cell(flexcpText, r, 0)
    Next
    
    If SUM��Ŀ > Val(fgMain.Cell(flexcpText, r - 1, 1)) Then
        SUM��Ŀ = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1)))
    ElseIf Abs(SUM��Ŀ) < 0.001 Then
        SUM��Ŀ = 0#
    Else
        SUM��Ŀ = Abs(SUM��Ŀ)
    End If
    If Right(lbl����, 3) = "�۷���" Then
        SUM��Ŀ = Abs(Val(fgMain.Cell(flexcpText, r - 1, 1))) - SUM��Ŀ
    End If
    SUM�ܷ� = SUM�ܷ� + SUM��Ŀ
    SUM��Ŀ = 0#

    Get���� = SUM�ܷ�
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� ȡ������ͳ�Ƶĵȼ�
'==============================================================================
Private Function Get�ȼ�() As String
    Dim �ȼ�1           As Long         '�ף�3
    Dim �ȼ�2           As Long         '�ң�2
    Dim �ȼ�            As Long         '����1
    Dim ����            As Single
    Dim ��ֵ            As Single
    Dim ��ֵ            As Single
    Dim r               As Long
    
    On Error GoTo errH
    
    ���� = Val(lbl�÷�)
    ��ֵ = Val(Mid(lbl��ֵ, 4))
    ��ֵ = Val(Mid(lbl��ֵ, 4))
    If ���� < ��ֵ Then
        �ȼ�1 = 1
    ElseIf ���� < ��ֵ Then
        �ȼ�1 = 2
    Else
        �ȼ�1 = 3
    End If
    
    �ȼ�2 = 3
    For r = 1 To fgMain.Rows - 1
        If fgMain.Cell(flexcpText, r, 4) = "������" Then
            If fgMain.Cell(flexcpText, r, 10) = "��" Then
                Get�ȼ� = "���ϸ�"
                Exit Function
            ElseIf fgMain.Cell(flexcpText, r, 10) = "��" Then
                If �ȼ�2 > 2 Then �ȼ�2 = 2
            ElseIf fgMain.Cell(flexcpText, r, 10) = "��" Then
                If �ȼ�2 > 1 Then �ȼ�2 = 1
            End If
        ElseIf fgMain.Cell(flexcpText, r, 4) = "�Ҽ�" Then
            If �ȼ�2 > 2 Then �ȼ�2 = 2
        ElseIf fgMain.Cell(flexcpText, r, 4) = "����" Then
            If �ȼ�2 > 1 Then �ȼ�2 = 1
        End If
    Next
    
    'ȡ�ȼ�1��ȼ�2����Сֵ��
    If �ȼ�1 > �ȼ�2 Then
        �ȼ� = �ȼ�2
    Else
        �ȼ� = �ȼ�1
    End If
    
    Select Case �ȼ�
    Case 1
        Get�ȼ� = "��"
    Case 2
        Get�ȼ� = "��"
    Case 3
        Get�ȼ� = "��"
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� ������̿���
'==============================================================================
Private Sub fgMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    
    If KeyCode = vbKeyEscape Then
        CmdCancel_Click
    End If
    If KeyCode = vbKeyReturn Then
        If Shift = 2 Then
            CmdOK_Click
            Exit Sub
        End If
        Select Case fgMain.Cell(flexcpText, fgMain.Row, 3)
            Case "�׼�", "�Ҽ�", "����", "������"
                fgMain_StartEdit fgMain.Row, fgMain.Col, False
            Case Else
                KeyCode = 0
                If fgMain.Row < fgMain.Rows - 1 Then
                    fgMain.Row = fgMain.Row + 1
                    If fgMain.Row < fgMain.Rows - 3 Then
                        fgMain.ShowCell fgMain.Row + 2, 4
                    Else
                        fgMain.ShowCell fgMain.Row, 4
                    End If
                Else
                    If cmdOK.Enabled Then cmdOK.SetFocus
                End If
        End Select
    ElseIf KeyCode = vbKeyDelete Then
        fgMain.Cell(flexcpText, fgMain.Row, 4) = ""
    End If
    edKey = KeyCode
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������̿���
'==============================================================================
Private Sub fgMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errH
    
    If Col = 9 Then Exit Sub
        
    Cancel = True
    edRow = fgMain.Row
    edCol = fgMain.Col
    If edCol = 5 Then Exit Sub
    
    Select Case fgMain.Cell(flexcpText, fgMain.Row, 3)
        Case "�׼�", "�Ҽ�", "����", "������"
            '�б�̬�ı�
            lst����.Clear
            lst����.AddItem "1 - ��"
            lst����.AddItem "2 - " + fgMain.Cell(flexcpText, fgMain.Row, 3)
            txt����.Text = ""
        
            txt����.Visible = False
            Select Case fgMain.Cell(flexcpText, fgMain.Row, 4)
                Case "�׼�", "�Ҽ�", "����", "������"
                    lst����.ListIndex = 1
                Case Else
                    lst����.ListIndex = 0
            End Select
            lst����.Move fgMain.CellLeft, fgMain.CellTop + fgMain.CellHeight, fgMain.CellWidth
            lst����.Visible = True
            lst����.SetFocus
        Case Else
            txt����.Move fgMain.CellLeft, fgMain.CellTop, fgMain.CellWidth - Screen.TwipsPerPixelX, fgMain.CellHeight - Screen.TwipsPerPixelY
            txt����.Text = fgMain.Text
            If edKey >= 96 And edKey <= 105 Then 'С����
                txt����.Text = edKey - 96
                txt����.SelStart = 1
            ElseIf edKey = vbKeyDecimal Then
                txt����.Text = "."
                txt����.SelStart = 1
            ElseIf edKey > 32 Then
                txt����.Text = Chr(edKey)
                txt����.SelStart = 1
            ElseIf edKey = 32 Then
                txt����.SelStart = 0
                txt����.SelStart = 32000
            Else
                txt����.SelStart = 0
                txt����.SelLength = 32000
            End If
            
            txt����.Visible = True
            txt����.SetFocus
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �ո������
'==============================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    
    If KeyAscii = vbKeyEscape Then
        CmdCancel_Click
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���λ�ñ仯����
'==============================================================================
Private Sub picLeft_Resize()
On Error Resume Next
    pic������Ϣ.Move 135, 135
    pic������Ϣ.Move 135, pic������Ϣ.Top + pic������Ϣ.Height + 135
    pic��Ŀ��Ϣ.Move 135, pic������Ϣ.Top + pic������Ϣ.Height + 135, pic��Ŀ��Ϣ.Width, IIf(imgXMXX.Tag <> "", pic��Ŀ��Ϣ.Height, Abs(picLeft.ScaleHeight - pic������Ϣ.Height - pic������Ϣ.Height - 135 * 4))
    pic������Ϣ.Cls
    pic������Ϣ.PaintPicture imgBGBlue.Picture, 0, 0, pic������Ϣ.Width, 360, 0, 0, imgBGBlue.Width, 360
    pic������Ϣ.PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    pic������Ϣ.PaintPicture imgBGBlue.Picture, pic������Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
    pic������Ϣ.PaintPicture imgBGBlue.Picture, 0, pic������Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic������Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    
    pic������Ϣ.Cls
    pic������Ϣ.PaintPicture imgBG.Picture, 0, 0, pic������Ϣ.Width, 360, 0, 0, imgBG.Width, 360
    pic������Ϣ.PaintPicture imgBG.Picture, 0, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic������Ϣ.PaintPicture imgBG.Picture, pic������Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, imgBG.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic������Ϣ.PaintPicture imgBG.Picture, 0, pic������Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic������Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBG.Height - Screen.TwipsPerPixelY, imgBG.Width, Screen.TwipsPerPixelY
    pic��Ŀ��Ϣ.Cls
    
    pic��Ŀ��Ϣ.PaintPicture imgBG.Picture, 0, 0, pic��Ŀ��Ϣ.Width, 360, 0, 0, imgBG.Width, 360
    pic��Ŀ��Ϣ.PaintPicture imgBG.Picture, 0, 360, Screen.TwipsPerPixelX, pic��Ŀ��Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic��Ŀ��Ϣ.PaintPicture imgBG.Picture, pic��Ŀ��Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic��Ŀ��Ϣ.Height - 360, imgBG.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
    pic��Ŀ��Ϣ.PaintPicture imgBG.Picture, 0, pic��Ŀ��Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic��Ŀ��Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBG.Height - Screen.TwipsPerPixelY, imgBG.Width, Screen.TwipsPerPixelY
    imgBRXX.Move pic������Ϣ.ScaleWidth - imgBRXX.Width - 100
    imgFAXX.Move pic������Ϣ.ScaleWidth - imgFAXX.Width - 100
    imgXMXX.Move pic��Ŀ��Ϣ.ScaleWidth - imgXMXX.Width - 100
    Refresh
End Sub

'==============================================================================
'=���ܣ� �Ҳ�λ�ñ仯����
'==============================================================================
Private Sub picRight_Resize()
On Error Resume Next
    With fgMain
        .Height = picRight.ScaleHeight - .Top
        .Width = picRight.ScaleWidth - .Left
    End With
End Sub

'==============================================================================
'=���ܣ� �Ҳ�λ�ñ仯����
'==============================================================================
Private Sub pic��Ŀ��Ϣ_Resize()
On Error Resume Next
    txt��Ŀ��Ϣ.Move txt��Ŀ��Ϣ.Left, txt��Ŀ��Ϣ.Top, txt��Ŀ��Ϣ.Width, Abs(pic��Ŀ��Ϣ.ScaleHeight - txt��Ŀ��Ϣ.Top - 135)
End Sub

'==============================================================================
'=���ܣ� ֵ�޸ĺ���
'==============================================================================
Private Sub txtNo_Change()
    On Error GoTo errH
    
    txtNo.Tag = "Changed"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ֵ�޸ĺ���
'==============================================================================
Private Sub txtNo_GotFocus()
    On Error GoTo errH
    
    zlControl.TxtSelAll txtNo
    ShowTips picRight, txtNo, "��A�򣭿�ͷ������:       ����ID" & vbCrLf & _
        "��B�򣫿�ͷ������:       סԺ��" & vbCrLf & _
        "��C�򣯿�ͷ������:       ��λ��" & vbCrLf & _
        "��D�򣪿�ͷ������:       �����" & vbCrLf & _
        "������:                             ���￨��" & vbCrLf & _
        "�����������Ϊ������������", "���ٶ�λʹ�ü���" & vbCrLf
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����¼�밴������
'==============================================================================
Private Sub txtNo_KeyPress(KeyAscii As Integer)

    Dim StrText         As String
    Dim strTmp          As String
    Dim bytFilterMode   As Byte
    Dim lng����ID       As Long
    Dim blnCard         As Boolean
    
    On Error GoTo errH
    
    If txtNo.Tag <> "" Then
        '���￨��

        blnCard = zlCommFun.InputIsCard(txtNo, KeyAscii, ParamInfo.ϵͳ��)
        If blnCard Then
            If Len(txtNo.Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txtNo.Text <> "" Then
                If KeyAscii <> 13 Then
                    txtNo.Text = txtNo.Text & Chr(KeyAscii)
                    txtNo.SelStart = Len(txtNo.Text)
                    KeyAscii = 0
                End If
                
                StrText = txtNo.Text
                bytFilterMode = 1
            End If
        End If
    End If
    
    Select Case KeyAscii
        '------------------------------------------------------------------------------------------------------------------
        Case vbKeyReturn
            KeyAscii = 0
            If txtNo.Tag = "Changed" Then
                If InStr(txtNo.Text, "'") Then
                    ShowSimpleMsg "¼����������зǷ��ַ� ' ��"
                    Exit Sub
                End If
                StrText = txtNo.Text
                Select Case UCase(Left(StrText, 1))
                    Case "-", "A"                 '����id
                        bytFilterMode = 2
                        StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                    Case "+", "B"                 'סԺ��
                        bytFilterMode = 3
                        StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                    Case "*", "D"                 '�����
                        bytFilterMode = 4
                        StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                    Case Else                     '����
                        txtNo.Tag = ""
                        zlCommFun.PressKey vbKeyTab
                        Exit Sub
                End Select
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        '------------------------------------------------------------------------------------------------------------------
        Case vbKeyEscape
            Call CmdCancel_Click
        '------------------------------------------------------------------------------------------------------------------
        Case Else
            If Chr(KeyAscii) = "'" Then KeyAscii = 0
            If Chr(KeyAscii) = "|" Then KeyAscii = 0
    End Select
    
    If StrText <> "" And bytFilterMode > 0 Then
        Call ���Ҳ���(StrText, bytFilterMode)
        txtNo.Tag = ""
    End If
    
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

'==============================================================================
'=���ܣ� �������Ƿ���ֵ������ʾ��
'==============================================================================
Private Sub txt����_Change()
    Dim Num As Single
    On Error GoTo errH
    
    Num = Abs(Val(txt����.Text))
    If Num > 9999 Then
        txt����.Text = 9999
    ElseIf InStr(1, fgMain.Cell(flexcpText, fgMain.Row, 3), "/") > 0 Then
        
    ElseIf Num > Val(fgMain.Cell(flexcpText, fgMain.Row, 3)) Then
        txt����.Text = Val(fgMain.Cell(flexcpText, fgMain.Row, 3))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������Ƿ���ֵ������ʾ��
'==============================================================================
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    Select Case KeyCode
        Case vbKeyRight
            If txt����.SelStart = Len(txt����.Text) Then
                fgMain.SetFocus
                fgMain.Col = fgMain.Col + 1
            End If
        Case vbKeyLeft
            If txt����.SelStart = Len(txt����.Text) Then
                fgMain.SetFocus
                fgMain.Col = fgMain.Col - 1
            End If
        Case vbKeyUp
            fgMain.SetFocus
            fgMain.Row = fgMain.Row - 1
            fgMain.ShowCell fgMain.Row, 4
        Case vbKeyReturn, vbKeyDown
            '�������Ƿ���ֵ������ʾ��
            If Len(txt����) > 0 And IsNumeric(txt����) = False Or Abs(Val(txt����)) > 9999 Then
                ShowTips picRight, txt����, "��������ȷ������", "��ʽ����", 2000, fgMain.Top
                txt����.SelStart = 0
                txt����.SelLength = Len(txt����)
                txt����.SetFocus
                Exit Sub
            End If
            
            fgMain.SetFocus
            If fgMain.Row <> fgMain.Rows - 1 Then
                fgMain.Row = fgMain.Row + 1
                fgMain.ShowCell fgMain.Row, 4
            Else
                If cmdOK.Enabled Then cmdOK.SetFocus
            End If
        Case vbKeyEscape
            KeyCode = 0
            txt����.Text = fgMain.TextMatrix(edRow, edCol)
            fgMain.SetFocus
    End Select
    edKey = KeyCode
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ְ�������
'==============================================================================
Private Sub lst����_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    
    Select Case KeyCode
        Case vbKeyReturn
            fgMain.SetFocus
            If fgMain.Row = fgMain.Rows - 1 Then
                If cmdOK.Enabled Then cmdOK.SetFocus
                Exit Sub
            End If
            fgMain.Row = fgMain.Row + 1
            fgMain.ShowCell fgMain.Row, 4
        Case vbKeyEscape
            KeyCode = 0
            Select Case fgMain.TextMatrix(edRow, edCol)
                Case "�׼�", "�Ҽ�", "����", "������"
                    lst����.ListIndex = 1
                Case Else
                    lst����.ListIndex = 0
            End Select
            fgMain.SetFocus
        Case vbKey1, vbKeyNumpad1
            lst����.ListIndex = 0
            fgMain.SetFocus
            fgMain.Row = fgMain.Row + 1
            fgMain.ShowCell fgMain.Row, 4
        Case vbKey2, vbKeyNumpad2
            lst����.ListIndex = 1
            fgMain.SetFocus
            If fgMain.Row = fgMain.Rows - 1 Then
                If cmdOK.Enabled Then cmdOK.SetFocus
                Exit Sub
            End If
            fgMain.Row = fgMain.Row + 1
            fgMain.ShowCell fgMain.Row, 4
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ְ�������
'==============================================================================
Private Sub txt����_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    Select Case KeyAscii
        Case 13, 27:
            KeyAscii = 0
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������ݼ����Ч��
'==============================================================================
Private Sub txt����_LostFocus()
    Dim Num         As Single
    
    On Error GoTo errH
    
    Num = Abs(Val(txt����.Text))
    If Num > 9999 Then
        txt���� = 9999
        Exit Sub
    End If
    If Num < 0.01 Then
        fgMain.TextMatrix(edRow, edCol) = ""
    Else
        If Num < 1 Then
            fgMain.TextMatrix(edRow, edCol) = Format(Num, "0.0")
        Else
            fgMain.TextMatrix(edRow, edCol) = Num
        End If
    End If
    
    txt����.Visible = False
    edKey = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʾ��ʾ
'==============================================================================
Private Sub ShowTips(ctl0 As Control, ctl As Control, str���� As String, Optional str���� As String = "��ʾ��Ϣ", Optional lngʱ�� As Long = 4500, Optional �����߶� As Long = 0)
    Dim X       As Single
    Dim Y       As Single
    
    On Error GoTo errH
    
    X = (ctl.Left + ctl.Width / 2 + ctl0.Left) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height + ctl0.Top + �����߶�) / Screen.TwipsPerPixelY
    If Len(str����) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        
        tipPopup1.TimeOut = lngʱ��
        tipPopup1.Title = str����
        tipPopup1.Text = str����
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������ĿID��ʾ��Ŀ����Ҫ��
'==============================================================================
Private Sub Show����Ҫ��(lngID As Long, ��Ŀ As String, ��׼��ֵ As String)
    Dim rs              As ADODB.Recordset
    On Error GoTo errH
    
    gstrSQL = "select ID,���� as ����Ҫ��,�ϼ�ID from �������ֱ�׼ Where ID= [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If Not rs.EOF Then
        If m_lngOldSJID > 0 And m_lngOldSJID = lngID Then Exit Sub
        If IsNull(rs.Fields("����Ҫ��")) Then
            txt��Ŀ��Ϣ = "���ƣ�" + ��Ŀ + "  " + IIf(Len(Trim(��׼��ֵ)) = 0, "", "(" + ��׼��ֵ + ")")
            txt��Ŀ��Ϣ = txt��Ŀ��Ϣ + vbCrLf
        Else
            If Len(rs.Fields("����Ҫ��")) > 0 Then
                txt��Ŀ��Ϣ = "���ƣ�" + ��Ŀ + "  " + IIf(Len(Trim(��׼��ֵ)) = 0, "", "(" + ��׼��ֵ + ")")
                txt��Ŀ��Ϣ = txt��Ŀ��Ϣ + vbCrLf + rs.Fields("����Ҫ��")
            End If
        End If
    Else
        txt��Ŀ��Ϣ = "��":
    End If
    m_lngOldSJID = m_lngCurSJID
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������ѯ
'==============================================================================
Private Sub ���Ҳ���(strID As String, ByVal bytFilterMode As Byte)
    Dim lngBRID         As Long
    Dim lngZYID         As Long
    Dim strSQL          As String
    Dim lngFAID         As Long
    Dim i               As Long
    Dim rs              As ADODB.Recordset
    Dim lngCurRowTMP    As Long
    Dim blnFinded       As Boolean
    
    On Error GoTo errH
    
    
    Select Case bytFilterMode
        Case 1              '���￨��
            
            strSQL = _
                "Select A.����ID,B.��ҳID " & _
                " From ������Ϣ A,������ҳ B " & _
                " Where A.����ID=B.����ID " & _
                " And Nvl(B.��ҳID,0)<>0 " & _
                " And A.���￨��=[1]"
                
        Case 2              '����ID
            strSQL = _
                "Select A.����ID,B.��ҳID " & _
                " From ������Ϣ A,������ҳ B " & _
                " Where A.����ID=B.����ID " & _
                " And Nvl(B.��ҳID,0)<>0 " & _
                " And A.����ID=[1]"
        Case 3              'סԺ��
            strSQL = _
                "Select A.����ID,B.��ҳID " & _
                " From ������Ϣ A,������ҳ B " & _
                " Where A.����ID=B.����ID " & _
                " And Nvl(B.��ҳID,0)<>0 " & _
                " And A.סԺ��=[1]"
        Case 4              '�����
            strSQL = _
                "Select A.����ID,B.��ҳID " & _
                " From ������Ϣ A,������ҳ B " & _
                " Where A.����ID=B.����ID " & _
                " And Nvl(B.��ҳID,0)<>0 " & _
                " And A.�����=[1]"
        Case Else            '����
            strSQL = _
                "Select A.����ID,B.��ҳID " & _
                " From ������Ϣ A,������ҳ B " & _
                " Where A.����ID=B.����ID " & _
                " And Nvl(B.��ҳID,0)<>0 " & _
                " And Upper(A.����)=[2]"
                
            strID = UCase(strID)
    End Select

    If mbln��Ŀ������ Then
        gstrSQL = strSQL & " and ��Ŀ���� is not null"
    Else
        gstrSQL = strSQL
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID), strID)
    If Not rs.EOF Then '�ҵ���¼
        lngBRID = rs("����ID")
'        lngZYID = Rs("��ҳID")
        If lngBRID <= 0 Then Exit Sub
        '��λ�����没����¼
        With frm��������.fg����_S
            lngCurRowTMP = .Row
            For i = lngCurRowTMP + 1 To .Rows - 1
                If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                    .Row = i
                    .ShowCell i, 2
                    blnFinded = True
                    Exit For
                End If
            Next
            If blnFinded = False Then '�����ǰ������û��ƥ�����ӵ�һ�п�ʼ���²�ѯ��
                For i = 1 To lngCurRowTMP
                    If Val(.Cell(flexcpText, i, 3)) = lngBRID Then
                        .Row = i
                        .ShowCell i, 2
                        blnFinded = True
                        Exit For
                    End If
                Next
            End If
        End With
        
        '���벡����Ϣ
        rs.Close
        If InStr(gstrPrivs, "���п���") = 0 Then    '�����п��ҹ���
            gstrSQL = "select ����,�Ա�,����ID,��ҳID,סԺ��,��Ժ����,��Ժ����,��Ժ����,��Ժ����,����ҽʦ,���λ�ʿ,סԺҽʦ,��Ŀ����,���ID,����ID,�ܷ�,�ȼ�,������,����ʱ��,�����,���ʱ��,�����޸�,��ע " & _
                      "from ��������������ͼ where ����ʱ�� is null and ���ʱ�� is null and ��Ժ���� = [1] and ����ID = [2]"
        Else
            gstrSQL = "select ����,�Ա�,����ID,��ҳID,סԺ��,��Ժ����,��Ժ����,��Ժ����,��Ժ����,����ҽʦ,���λ�ʿ,סԺҽʦ,��Ŀ����,���ID,����ID,�ܷ�,�ȼ�,������,����ʱ��,�����,���ʱ��,�����޸�,��ע " & _
                      "from ��������������ͼ where ����ʱ�� is null and ���ʱ�� is null and ����ID=[2]"
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDeptName, lngBRID)

        If Not rs.EOF Then
            lngFAID = NVL(rs("����ID"), 0)
            lngZYID = rs("��ҳID")
            
            m_bln��α༭ = True
            cmdOK.Enabled = True
            ShowForm "����", 0, lngBRID, lngZYID, lngFAID, m_lng����ID
            Exit Sub
        Else
            cmdOK.Enabled = False
        End If
    Else
        cmdOK.Enabled = False
    End If
    MsgBox "û���ҵ�ָ�����������߸ò����Ѿ����֣����������롣", vbExclamation, gstrSysName
    Call ClearResults
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������������ѯ
'==============================================================================
Private Sub ClearResults()
    Dim i As Long
    
    On Error GoTo errH
    
    lbl����.Caption = "��   ��:"
    lblסԺ��.Caption = "ס Ժ ��:"
    lblסԺ����.Caption = "סԺ����:"
    lbl��Ժ����.Caption = "��Ժ����: " & gstrDeptName
    lblסԺҽʦ.Caption = "סԺҽʦ:"
    lbl��Ŀ����.Caption = "��Ŀ����:"
    chk�����޸�.Value = vbUnchecked
    txtNo.Text = ""
    txt��ע.Text = ""
    
    For i = 1 To fgMain.Rows - 1                '���ԭ�����ֽ��
        fgMain.Cell(flexcpText, i, 4) = ""
        fgMain.Cell(flexcpText, i, 5) = ""
        fgMain.Cell(flexcpText, i, 9) = ""
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



