VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.8#0"; "zlIDKind.ocx"
Begin VB.Form frmPatiFamily 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ϵ"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatiFamily.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6975
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5760
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4560
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame fraSplit1 
      BackColor       =   &H00C0C0C0&
      Height          =   45
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Width           =   8055
   End
   Begin VB.Frame fraPatiInfo 
      BorderStyle     =   0  'None
      Caption         =   "������Ϣ"
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   6975
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4920
         TabIndex        =   28
         Tag             =   "�Ա�:"
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2760
         TabIndex        =   27
         Tag             =   "�Ա�:"
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   5280
         TabIndex        =   26
         Tag             =   "�Ա�:"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2940
         TabIndex        =   25
         Tag             =   "�Ա�:"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Tag             =   "����:"
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   300
         TabIndex        =   23
         Tag             =   "����:"
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblJZK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10101010101"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3480
         TabIndex        =   22
         Tag             =   "��������:"
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "201502101"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   840
         TabIndex        =   21
         Tag             =   "����:"
         Top             =   480
         Width           =   810
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͨ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   5760
         TabIndex        =   12
         Tag             =   "��������:"
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "30��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   5760
         TabIndex        =   11
         Tag             =   "����:"
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ֪"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3480
         TabIndex        =   10
         Tag             =   "�Ա�:"
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����༪"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   840
         TabIndex        =   9
         Tag             =   "����:"
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame fraSplit2 
      BackColor       =   &H00C0C0C0&
      Height          =   45
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   8055
   End
   Begin VB.Frame fraGroup 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   19
      Top             =   1920
      Width           =   6975
      Begin VSFlex8Ctl.VSFlexGrid vsfamily 
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6750
         _cx             =   11906
         _cy             =   2514
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiFamily.frx":6852
         ScrollTrack     =   0   'False
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
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picdel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   6240
            Picture         =   "frmPatiFamily.frx":6913
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.Frame fraPatiCard 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "ˢ�������������,��Enter����ȡ������Ϣ"
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtPatiPWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin zlIDKind.PatiIdentify PatiIdentifyPati 
         Height          =   375
         Left            =   600
         TabIndex        =   0
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPatiFamily.frx":D165
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   2
         InputAppearance =   2
         ShowSortName    =   -1  'True
         DefaultCardType =   "0"
         IDkindBorderStyle=   1
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CardNOForColor  =   -2147483635
         MustBrushCard   =   -1  'True
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
         BackColor       =   16777215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4500
         TabIndex        =   15
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   360
      End
   End
   Begin VB.Frame fraFamilyCard 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   16
      ToolTipText     =   "ˢ�������������,��Enter��¼�������Ϣ"
      Top             =   1440
      Width           =   6975
      Begin VB.TextBox txtFamilyPWD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4950
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
      Begin zlIDKind.PatiIdentify PatiIdentifyFamily 
         Height          =   375
         Left            =   630
         TabIndex        =   2
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPatiFamily.frx":D22C
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   2
         InputAppearance =   2
         ShowSortName    =   -1  'True
         DefaultCardType =   "0"
         IDkindBorderStyle=   1
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CardNOForColor  =   -2147483635
         MustBrushCard   =   -1  'True
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.Label lblPWD 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4500
         TabIndex        =   18
         Top             =   90
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   90
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmPatiFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytFunc As Long    '1-�鿴;2-�༭
Private mbytCount As Byte  '��¼����������
Private mlng����ID As Long
Private mlngModule As Long
Private mobjKeyboard As Object
Private mblnReturn As Boolean
Private msinTime As Single

Private Type T_Pati
    ����ID As Long
    ���� As String
    �Ա� As String
    ���� As String
    ���￨�� As String
    ���� As String
End Type

Private mPati As T_Pati
Private mFamily As T_Pati

Private Const C_FamilyColumHeader = "��ϵ,1505,4;����,1370,4;�Ա�,705,4;����,705,4;���￨��,1545,4;��ϸ,495,4; ����,300,4" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_COLOR_��ɫ = &H80000000
Private Const C_COLOR_��ɫ = &H80000005

Public Sub ShowMe(ByRef frmMain As Object, ByVal lng����ID As Long, ByVal bytFunc As Byte, ByVal lngModul As Long)
'����:��ʾ������
'����: objFrmMain-������
'       =1-�鿴,2-�༭
'      lng����ID-�鿴ʱ���� ��bytFunc=1ʱ����,bytFunc=2ʱˢ����ȡ��
'      lngModul ģ���
'     str��ϵ-�����������ı���ʾ
'     rsFamily-���ڻ��没�˼���
    mbytFunc = bytFunc
    mlngModule = lngModul
    If mbytFunc = 1 Then
        mlng����ID = lng����ID
        Me.Caption = "������Ϣ"
    Else
        mlng����ID = 0
        Me.Caption = "�����Ǽ�"
    End If
    Me.Show 1, frmMain
End Sub

Private Sub cmdCancel_Click()
    Dim i As Long
    
    With vsfamily
        For i = 1 To .Rows - 1
            If InStr(",2,3,4,", "," & .RowData(i) & ",") > 0 Then
                If MsgBox("���ڱ༭��δ�������Ϣ����ȷ��Ҫȡ����", vbOKCancel + vbQuestion + vbDefaultButton1, gstrSysName) = vbCancel Then
                    Exit Sub
                End If
            End If
        Next
    End With
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not SavePatiFamily Then Exit Sub
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mblnReturn Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0 '����ˢ���Դ��س�
            mblnReturn = False
        End If
    End If
End Sub

Private Sub Form_Load()
    '��ʼ��
    
    Call ClearPatiInfo
    Call InitVsFamily
    Call LoadPatiFamily
    picdel.Visible = False
    If mbytFunc = 1 Then
        '�鿴
        cmdCancel.Caption = "�ر�(&C)"
        Call LoadPati
    ElseIf mbytFunc = 2 Then
        '�༭
        Call CreateSquareCardObject(Me, mlngModule)
        If Not gobjSquare Is Nothing Then
            PatiIdentifyPati.MustBrushCard = True   '����ˢ��
            PatiIdentifyPati.OnlyThreeCard = True
            Call PatiIdentifyPati.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
            
            PatiIdentifyFamily.MustBrushCard = True  '����ˢ��
            PatiIdentifyFamily.OnlyThreeCard = True
            Call PatiIdentifyFamily.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
        End If
        '������Ӽ���
        Call CreateObjectKeyboard
        PatiIdentifyFamily.Enabled = False
        txtFamilyPWD.Enabled = False
        PatiIdentifyFamily.BackColor = C_COLOR_��ɫ
        txtFamilyPWD.BackColor = C_COLOR_��ɫ
        cmdCancel.Caption = "ȡ��(&C)"
    End If
    
End Sub

Private Sub Form_Resize()
    Dim lngW As Long
    
    If mbytFunc = 1 Then  '�鿴
        lngW = 6975
        Me.Width = 7065: Me.Height = 3705
        fraPatiCard.Visible = False
        fraPatiInfo.Move 0, 0, lngW, 735
        fraSplit1.Move 0, fraPatiInfo.Top + fraPatiInfo.Height + 45, lngW, 45
        fraFamilyCard.Visible = False
        fraGroup.Move 0, fraSplit1.Top + fraSplit1.Height + 120, lngW, 1575
        fraSplit2.Move 0, fraGroup.Top + fraGroup.Height + 120, lngW, 45
        cmdCancel.Move 5760, 2805, 1095, 350
    ElseIf mbytFunc = 2 Then
        lngW = 6975
        Me.Width = 7065: Me.Height = 4740
        fraPatiCard.Visible = True
        fraFamilyCard.Visible = True
        fraPatiCard.Move 0, 0, lngW, 495
        fraPatiInfo.Move 0, fraPatiCard.Height, lngW, 735
        fraSplit1.Move 0, fraPatiInfo.Top + fraPatiInfo.Height + 45, lngW, 45
        fraFamilyCard.Move 0, fraSplit1.Top + fraSplit1.Height + 120, lngW, 375
        fraGroup.Move 0, fraFamilyCard.Top + fraFamilyCard.Height + 120, lngW, 1575
        fraSplit2.Move 0, fraGroup.Top + fraGroup.Height + 120, lngW, 45
        cmdOK.Move 4560, 3840, 1095, 350
        cmdCancel.Move 5760, 3840, 1095, 350
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjKeyboard = Nothing
End Sub

Private Sub PatiIdentifyFamily_Change()
    If Trim(PatiIdentifyFamily.Text) = "" Then
        txtFamilyPWD.Text = ""
        Call ReSetPati(mFamily)
    End If
End Sub

Private Sub PatiIdentifyFamily_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    If objHisPati Is Nothing Then
        mFamily.����ID = 0
        blnCancel = True
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
        
        MsgBox "������Ϣδ�ҵ�������ԭ��:" & vbCrLf & _
                Space(4) & "(1)��ǰѡ��Ŀ����͡�" & PatiIdentifyFamily.GetCurCard.���� & "�������ֿ������Ͳ�����" & vbCrLf & _
                Space(4) & "(2)���ֿ�δ�󶨲�����Ϣ��", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    Else
        mbytCount = 3
        mFamily.����ID = Val(objHisPati.����ID)
        mFamily.���� = objHisPati.����
        mFamily.���� = objHisPati.����
        mFamily.�Ա� = objHisPati.�Ա�
        mFamily.���￨�� = objHisPati.���￨��
        mFamily.���� = objHisPati.����
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
        txtFamilyPWD.SetFocus
    End If
    
End Sub

Private Sub PatiIdentifyFamily_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If Val(PatiIdentifyFamily.Tag) <> Index Then
        PatiIdentifyFamily.Tag = Index
        PatiIdentifyFamily.Text = ""
    End If
End Sub

Private Sub PatiIdentifyFamily_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(PatiIdentifyFamily.Text) = PatiIdentifyFamily.GetCurCard.���ų��� - 1 And PatiIdentifyFamily.SelLength <> Len(PatiIdentifyFamily.Text)
    If KeyAscii = vbKeyReturn Or blnCard Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            txtFamilyPWD.SetFocus
        End If
    End If
End Sub

Private Sub PatiIdentifyFamily_LostFocus()
    If mFamily.����ID = 0 And Len(PatiIdentifyFamily.Text) <> 0 Then
        MsgBox "������Ϣδ�ҵ�������ԭ��:" & vbCrLf & _
                Space(4) & "(1)��ǰѡ��Ŀ����͡�" & PatiIdentifyFamily.GetCurCard.���� & "�������ֿ������Ͳ�����" & vbCrLf & _
                Space(4) & "(2)���ֿ�δ�󶨲�����Ϣ��", vbInformation + vbOKOnly, gstrSysName
        PatiIdentifyFamily.SetFocus
    End If
End Sub

Private Sub PatiIdentifyPati_Change()
    If Trim(PatiIdentifyPati.Text) = "" Then
        mlng����ID = 0
        mbytCount = 3
        txtPatiPWD.Text = ""
        Call ClearPatiInfo
        Call ReSetPati(mPati)
    End If
End Sub

Private Sub PatiIdentifyPati_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    
    If objHisPati Is Nothing Then
        blnCancel = True
        mlng����ID = 0  '���δ�ҵ�����
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
        
        MsgBox "������Ϣδ�ҵ�������ԭ��:" & vbCrLf & _
            Space(4) & "(1)��ǰѡ��Ŀ����͡�" & PatiIdentifyPati.GetCurCard.���� & "�������ֿ������Ͳ�����" & vbCrLf & _
            Space(4) & "(2)���ֿ�δ�󶨲�����Ϣ��", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    Else
        Debug.Print "1"
        mbytCount = 3
        mlng����ID = objHisPati.����ID
        mPati.����ID = objHisPati.����ID
        mPati.���� = objHisPati.����
        mPati.���� = objHisPati.����
        mPati.���� = objHisPati.����
        mPati.�Ա� = objHisPati.�Ա�
        
        txtPatiPWD.SetFocus
        mblnReturn = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnReturn Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnReturn = False
    End If
    
End Sub

Private Sub PatiIdentifyPati_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If Val(PatiIdentifyPati.Tag) <> Index Then
        PatiIdentifyPati.Tag = Index
        PatiIdentifyPati.Text = ""
    End If
End Sub

Private Sub PatiIdentifyPati_KeyPress(KeyAscii As Integer)
     Dim blnCard As Boolean
    '�Ƿ�ˢ�����
    mblnReturn = False
    blnCard = KeyAscii <> 8 And Len(PatiIdentifyPati.Text) = PatiIdentifyPati.GetCurCard.���ų��� - 1 And PatiIdentifyPati.SelLength <> Len(PatiIdentifyPati.Text)
    If KeyAscii = vbKeyReturn Or blnCard Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            txtPatiPWD.SetFocus
        End If
    End If
End Sub

Private Sub PatiIdentifyPati_LostFocus()
    If mlng����ID = 0 And Len(PatiIdentifyPati.Text) <> 0 Then
        MsgBox "������Ϣδ�ҵ�������ԭ��:" & vbCrLf & _
                Space(4) & "(1)��ǰѡ��Ŀ����͡�" & PatiIdentifyPati.GetCurCard.���� & "�������ֿ������Ͳ�����" & vbCrLf & _
                Space(4) & "(2)���ֿ�δ�󶨲�����Ϣ��", vbInformation + vbOKOnly, gstrSysName
        PatiIdentifyPati.SetFocus
    End If
End Sub

Private Sub txtFamilyPWD_GotFocus()
    If PatiIdentifyFamily.Text = "" Then
        MsgBox "����ˢ����¼�����롣", vbInformation, gstrSysName
        If PatiIdentifyFamily.Enabled Then PatiIdentifyFamily.SetFocus
        Exit Sub
    ElseIf Val(mFamily.����ID) = 0 And PatiIdentifyFamily.Text <> "" Then
        On Error Resume Next
        PatiIdentifyFamily.SetFocus '���벿����,ִ�е��˴� PatiIdentifyPati.SetFocus��ᱨ��
        Err.Clear: On Error GoTo 0
        Exit Sub
    ElseIf mlng����ID <> 0 Then
        If mFamily.���� = "" Then Call txtFamilyPWD_KeyPress(vbKeyReturn): Exit Sub
    End If
End Sub

Private Sub txtFamilyPWD_KeyPress(KeyAscii As Integer)
    Dim strPassWord As String
    
    If KeyAscii = 22 Then
        KeyAscii = 0 '������ճ��
    ElseIf InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strPassWord = gobjCommFun.zlStringEncode(txtFamilyPWD.Text)
        If strPassWord <> mFamily.���� Then
            If mbytCount = 1 Then
                MsgBox "���������������,���������룡", vbExclamation, gstrSysName
            Else
                MsgBox "�����������", vbExclamation, gstrSysName
            End If
            txtFamilyPWD.Text = "": mbytCount = mbytCount - 1
            If mbytCount = 0 Then
                PatiIdentifyFamily.Text = ""
                PatiIdentifyFamily.SetFocus
            ElseIf txtFamilyPWD.Enabled Then
                txtFamilyPWD.SetFocus
            End If
            Exit Sub
        Else
            If ADDFamily Then
                If cmdOK.Enabled = False Then cmdOK.Enabled = True
            Else
                If PatiIdentifyFamily.Enabled Then PatiIdentifyFamily.SetFocus
            End If
            PatiIdentifyFamily.Text = ""
        End If
    End If
End Sub

Private Sub txtPatiPWD_GotFocus()
    If PatiIdentifyPati.Text = "" Then
        MsgBox "����ˢ����¼�����롣", vbInformation, gstrSysName
        PatiIdentifyPati.SetFocus
        Exit Sub
    ElseIf mlng����ID = 0 And PatiIdentifyPati.Text <> "" Then
        On Error Resume Next
        PatiIdentifyPati.SetFocus '���벿����,ִ�е��˴� PatiIdentifyPati.SetFocus��ᱨ��
        Err.Clear: On Error GoTo 0
        Exit Sub
    ElseIf mlng����ID <> 0 Then
        If mPati.���� = "" Then Call txtPatiPWD_KeyPress(vbKeyReturn): Exit Sub
    End If
    Call gobjControl.TxtSelAll(txtPatiPWD)
    Call OpenPassKeyboard(txtPatiPWD, False)
End Sub

Private Sub txtPatiPWD_KeyPress(KeyAscii As Integer)
    Dim strPassWord As String
    Dim intRet As Integer
    If KeyAscii = 22 Then
        KeyAscii = 0 '������ճ��
    ElseIf InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If mblnReturn Then
            mblnReturn = False: Exit Sub
        End If
        strPassWord = gobjCommFun.zlStringEncode(txtPatiPWD.Text)
        If strPassWord <> mPati.���� Then
            If mbytCount = 1 Then
                MsgBox "���������������,���������룡", vbExclamation, gstrSysName
            Else
                MsgBox "�����������", vbExclamation, gstrSysName
            End If
            txtPatiPWD.Text = "": mbytCount = mbytCount - 1
            If mbytCount = 0 Then
                Unload Me '������󣬿�����2��
            ElseIf txtPatiPWD.Enabled Then
                txtPatiPWD.SetFocus
            End If
            Exit Sub
        Else
            PatiIdentifyPati.Enabled = False
            txtPatiPWD.Enabled = False
            PatiIdentifyPati.BackColor = C_COLOR_��ɫ
            txtPatiPWD.BackColor = C_COLOR_��ɫ
            PatiIdentifyFamily.Enabled = True
            txtFamilyPWD.Enabled = True
            PatiIdentifyFamily.BackColor = C_COLOR_��ɫ
            txtFamilyPWD.BackColor = C_COLOR_��ɫ
            Call LoadPati
            Call LoadPatiFamily
            If PatiIdentifyFamily.Enabled Then PatiIdentifyFamily.SetFocus
        End If
    End If
End Sub

Private Sub txtPatiPWD_LostFocus()
    Call ClosePassKeyboard(txtPatiPWD)
End Sub

Private Sub vsfamily_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfamily
        If .Col = .ColIndex("��ϵ") Then
            If .TextMatrix(Row, Col) <> Nvl(.Cell(flexcpData, Row, Col)) And CByte(Nvl(.RowData(Row))) = 1 Then
                .RowData(Row) = 3 '����
                If cmdOK.Enabled = False Then cmdOK.Enabled = True
            ElseIf .TextMatrix(Row, Col) = Nvl(.Cell(flexcpData, Row, Col)) And CByte(Nvl(.RowData(Row))) = 3 Then
                .RowData(Row) = 1 'δ����
            End If
        End If
    End With
End Sub

Private Sub vsfamily_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfamily
        If Not Me.Visible Then Exit Sub
        If (OldRow <> NewRow Or OldRow = NewRow And OldRow = 1) And NewRow > .FixedRows - 1 Then
            If mbytFunc = 2 Then
                If Nvl(.Cell(flexcpData, NewRow, .ColIndex("����"))) = "" Then Exit Sub
                If Me.Visible Then
                    If picdel.Visible = False Then picdel.Visible = True
                End If
                picdel.Top = .Cell(flexcpTop, NewRow, .ColIndex("����"))
                picdel.Left = .Cell(flexcpLeft, NewRow, .ColIndex("����"))
            Else
                picdel.Visible = False
            End If
        End If
    End With
End Sub

Private Sub vsfamily_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfamily.ColIndex("��ϵ") = Col And mbytFunc = 2 And CLng(vsfamily.Cell(flexcpData, Row, vsfamily.ColIndex("����"))) <> 0 Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfamily_Click()
    With vsfamily
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If CLng(.Cell(flexcpData, .Row, .ColIndex("����"))) = 0 Then Exit Sub
        If .ColIndex("��ϸ") = .Col And .TextMatrix(.Row, .Col) = "��ϸ" Then
            frmDegreeCard.mlng����ID = CLng(.Cell(flexcpData, .Row, .ColIndex("����")))
            frmDegreeCard.mlng��ҳID = 0
            frmDegreeCard.Show 1, Me
        End If
    End With
End Sub

Private Sub vsfamily_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long

    With vsfamily
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow <= 0 Then Exit Sub
        If .ColData(.ColIndex("��ϸ")) > .Rows - 1 Then .ColData(.ColIndex("��ϸ")) = 0
        If lngCol = .ColIndex("��ϸ") And lngRow = .ColData(.ColIndex("��ϸ")) Then
            .Cell(flexcpFontUnderline, lngRow, lngCol) = True
            .ColData(lngCol) = lngRow
        ElseIf lngCol <> .ColIndex("��ϸ") Or .ColData(.ColIndex("��ϸ")) <> lngRow Then
            .Cell(flexcpFontUnderline, .ColData(.ColIndex("��ϸ")), .ColIndex("��ϸ")) = False
            .ColData(lngCol) = lngRow
        End If
    End With
End Sub

Private Sub InitVsFamily()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo ErrHand
    If mbytFunc = 2 Then
        strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ����ϵ Order by ����"
        Call gobjDatabase.OpenRecordset(rsTemp, strSQL, "����ϵ")
    
        With rsTemp
            Do While Not rsTemp.EOF
                strTmp = strTmp & "|" & Nvl(rsTemp!����)
            rsTemp.MoveNext
            Loop
        End With
        If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
    End If
    
    With vsfamily
        .Rows = 2
        gobjGrid.Init vsfamily, C_FamilyColumHeader
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionFree
        If mbytFunc = 1 Then
            .ColHidden(.ColIndex("����")) = True
        ElseIf strTmp <> "" And mbytFunc = 2 Then
            .ColComboList(.ColIndex("��ϵ")) = strTmp
        End If
    End With
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picdel_Click()
    Dim lngRow As Long
    Dim strSQL As String
    Dim i As Long
    Dim lngFlag As Long
    
    With vsfamily
        If MsgBox("��ȷ��Ҫɾ�����У�", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            If InStr(",1,3,", "," & .RowData(.Row) & ",") > 0 Then 'ԭʼ,�޸�
                '�������޷��ö���ɾ��,����ɾ��ʱ���ò������ݣ�����������
                .RowData(.Row) = 4    '��Ǽ�ɾ��
                .RowHidden(.Row) = True
                If cmdOK.Enabled = False Then cmdOK.Enabled = True
            Else
                .RemoveItem .Row
            End If
            
            picdel.Visible = False
            
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False Then
                    Exit For
                Else
                    lngFlag = lngFlag + 1
                End If
            Next
            If lngFlag = .Rows - 1 Then .Rows = .Rows + 1 'ȱʡ��ʾһ��
        End If
       
    End With
End Sub

Private Sub LoadPati()
'����:���ز�����Ϣ
    '���˼���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH

    strSQL = "Select a.�����, a.סԺ��, a.���￨��, a.����, a.�Ա�, a.����, a.�������� From ������Ϣ A Where a.����id = [1]"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "���˼���", mlng����ID)
    
    If rsTmp.RecordCount > 0 Then
        lblName.Caption = rsTmp!����
        lblAge.Caption = rsTmp!���� & ""
        lblSex.Caption = rsTmp!�Ա� & ""
        If rsTmp!סԺ�� & "" <> "" Then
            lblTag.Caption = "סԺ��:"
            lblNum.Caption = rsTmp!סԺ�� & ""
        Else
            lblTag.Caption = "�����:"
            lblNum.Caption = rsTmp!����� & ""
        End If
        lblJZK.Caption = rsTmp!���￨�� & ""
        lblPatiType.Caption = rsTmp!�������� & ""
    End If
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub


Private Sub LoadPatiFamily()
'����:���ز��˼�����Ϣ
    '���˼���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    If mlng����ID <> 0 Then
        strSQL = "Select a.����ID, a.��ϵ, b.���￨��, b.����, b.����, b.�Ա�,1 as ״̬ " & vbNewLine & _
                "From ���˼��� A, ������Ϣ B" & vbNewLine & _
                "Where a.����id = b.����id And a.����id = [1] And A.����ʱ�� IS NULL"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "���˼���", mlng����ID)
    End If
    With vsfamily
       .Rows = 2 'ȱʡ��ʾһ��
        If rsTmp Is Nothing Then Exit Sub
        For i = 1 To rsTmp.RecordCount
            .Rows = i + 1
            .TextMatrix(i, .ColIndex("��ϵ")) = rsTmp!��ϵ & ""
            .TextMatrix(i, .ColIndex("����")) = rsTmp!���� & ""
            .TextMatrix(i, .ColIndex("����")) = rsTmp!���� & ""
            .TextMatrix(i, .ColIndex("�Ա�")) = rsTmp!�Ա� & ""
            .TextMatrix(i, .ColIndex("���￨��")) = rsTmp!���￨�� & ""
            .TextMatrix(i, .ColIndex("��ϸ")) = "��ϸ"
            .RowData(i) = rsTmp!״̬ & "" '1-ԭʼ����
            
            .Cell(flexcpData, i, .ColIndex("��ϵ")) = rsTmp!��ϵ & ""
            .Cell(flexcpData, i, .ColIndex("����")) = rsTmp!����ID & ""
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("��ϸ")) = &HC00000
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Function SavePatiFamily() As Boolean
'����:���没�˼�����Ϣ
'
    Dim strSQL As String
    Dim i As Long
    Dim strDate As String
    Dim strDateDel As String
    Dim addDate As Date
    Dim blnSave As Boolean    '����Ƿ�����Ч����
    
    Dim colSQL As Collection
    
    On Error GoTo errH
    addDate = gobjDatabase.Currentdate
    strDate = Format(addDate, "YYYY-MM-DD HH:MM:SS")
    Set colSQL = New Collection

    With vsfamily
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("��ϵ")) = "" And CLng(.Cell(flexcpData, i, .ColIndex("����"))) <> 0 Then
                    MsgBox "�ò��˼�����" & .TextMatrix(i, .ColIndex("����")) & "���벡�ˡ�" & lblName.Caption & "���Ĺ�ϵδ¼��,����¼����ٱ��档", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                    .Row = i: .Col = .ColIndex("��ϵ")
                    .SetFocus
                    .ShowCell .Row, .Col
                    Exit Function
                ElseIf InStr(",2,3,4,", "," & .RowData(i) & ",") > 0 Then
                    blnSave = True   '���ڱ�����Ŀ
                End If
            Next
            
            If Not blnSave Then
                If MsgBox("��ǰδ�����κμ�����Ϣ���Ƿ��˳���", vbYesNo + vbInformation + vbDefaultButton1, gstrSysName) = vbYes Then
                    SavePatiFamily = True
                Else
                    SavePatiFamily = False
                End If
                Exit Function
            End If
            
            For i = 1 To .Rows - 1
                If CByte(Nvl(.RowData(i))) = 2 Then   '����
                    strSQL = " Zl_���˼���_Update(1," & mlng����ID & "," & .Cell(flexcpData, i, .ColIndex("����")) & ",'" & UserInfo.���� & _
                             "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),'" & .TextMatrix(i, .ColIndex("��ϵ")) & "')"          '����
                    colSQL.Add strSQL, "_" & colSQL.Count
                ElseIf CByte(Nvl(.RowData(i))) = 3 Then  '����
                    strSQL = " Zl_���˼���_Update(2," & mlng����ID & "," & .Cell(flexcpData, i, .ColIndex("����")) & ",'',NULL,'" & .TextMatrix(i, .ColIndex("��ϵ")) & "')"                             '����
                    colSQL.Add strSQL, "_" & colSQL.Count
                ElseIf CByte(Nvl(.RowData(i))) = 4 And .RowHidden(i) = True Then '��ɾ��
                    '���ɾ��ʱ����һ�����ѭ��ɾ��ʱΥ��ΨһԼ��
                    addDate = addDate + 1 / 24 / 60 / 60
                    strDateDel = "To_Date('" & Format(addDate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
                    strSQL = "Zl_���˼���_Update(3," & mlng����ID & "," & .Cell(flexcpData, i, .ColIndex("����")) & ",'',NULL,NULL,'" & UserInfo.���� & "'," & strDateDel & ")"
                    colSQL.Add strSQL, "_" & colSQL.Count
                End If
            Next
        End If
    End With
    
    '���������ύ
    For i = 1 To colSQL.Count
        Call gobjDatabase.ExecuteProcedure(CStr(colSQL(i)), "������Ϣ")
    Next
    
    SavePatiFamily = True
    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ADDFamily() As Boolean
    Dim i As Long
    
    With vsfamily
         '��������Ϊ���������Լ�
        If mFamily.����ID & "" = mlng����ID & "" Then
            MsgBox "���˼�����" & mFamily.���� & "���������Լ�,������¼�룡", vbInformation, gstrSysName
            Exit Function
        End If
            
        For i = .FixedRows To .Rows - 1
            '��� ͬһ�����˲�������¼��
            If mFamily.����ID & "" = .Cell(flexcpData, i, .ColIndex("����")) & "" Then
                If .RowHidden(i) = False Then
                    MsgBox "�ò��˼�����" & mFamily.���� & "���Ѿ�¼��,�������ظ�¼�룡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '���������Ѿ���ɾ��
        Next
        
        If .TextMatrix(.Rows - 1, .ColIndex("����")) <> "" Then .Rows = .Rows + 1
        .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = mFamily.����ID
        .TextMatrix(.Rows - 1, .ColIndex("����")) = mFamily.����
        .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = mFamily.�Ա�
        .TextMatrix(.Rows - 1, .ColIndex("����")) = mFamily.����
        .TextMatrix(.Rows - 1, .ColIndex("��ϸ")) = "��ϸ"
        .TextMatrix(.Rows - 1, .ColIndex("���￨��")) = mFamily.���￨��

        .RowData(.Rows - 1) = 2 '2-����
        .Cell(flexcpForeColor, .Rows - 1, .ColIndex("��ϸ")) = &HC00000
        .ShowCell .Rows - 1, .ColIndex("��ϵ") '��ʾ������
    End With
    ADDFamily = True
End Function

Private Sub ClearPatiInfo()

    lblName.Caption = ""
    lblAge.Caption = ""
    lblSex.Caption = ""
    lblTag.Caption = "סԺ��:"
    lblNum.Caption = ""
    lblJZK.Caption = ""
    lblPatiType.Caption = ""
    cmdOK.Enabled = False
End Sub

Private Sub ReSetPati(udtPati As T_Pati)
    udtPati.����ID = 0
    udtPati.���￨�� = ""
    udtPati.���� = ""
    udtPati.���� = ""
    udtPati.�Ա� = ""
    udtPati.���� = ""
End Sub
