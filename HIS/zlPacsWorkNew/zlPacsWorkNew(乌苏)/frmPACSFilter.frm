VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPACSFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPACSFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5460
      TabIndex        =   42
      Top             =   4500
      Width           =   5460
      Begin VB.ComboBox cboSchemaName 
         Height          =   330
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelSchema 
         Caption         =   "ɾ������(&D)"
         Height          =   375
         Left            =   2520
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveSchema 
         Caption         =   "���淽��(&S)"
         Height          =   375
         Left            =   3840
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkmore 
         Caption         =   "����(&M)"
         Height          =   375
         Left            =   1200
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "ȱʡ(&F)"
         Height          =   375
         Left            =   105
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   375
         Left            =   4200
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   375
         Left            =   3120
         TabIndex        =   37
         ToolTipText     =   "F2"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label labSchema 
         Caption         =   "��ѯ������"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picmore 
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
      Height          =   2745
      Left            =   -120
      ScaleHeight     =   2745
      ScaleWidth      =   5745
      TabIndex        =   46
      Top             =   4920
      Visible         =   0   'False
      Width           =   5745
      Begin VB.TextBox txt�������� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   31
         Top             =   405
         Width           =   3690
      End
      Begin VB.TextBox TxtӰ����� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   30
         Top             =   30
         Width           =   3690
      End
      Begin VB.Frame Frame1 
         Caption         =   "PACS�����ѯ"
         Height          =   1455
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   5250
         Begin VB.TextBox txtPacsRpt 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   3690
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   34
            Top             =   600
            Width           =   3690
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   35
            Top             =   960
            Width           =   3690
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "�������"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "�� ������"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "�� ����"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   1110
            Width           =   1095
         End
      End
      Begin VB.TextBox txt��� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   32
         Top             =   795
         Width           =   3690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   210
         Left            =   240
         TabIndex        =   55
         Top             =   465
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   210
         Left            =   240
         TabIndex        =   52
         Top             =   90
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   210
         Left            =   240
         TabIndex        =   51
         Top             =   855
         Width           =   420
      End
   End
   Begin VB.Frame Frabase 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   -120
      TabIndex        =   36
      Top             =   0
      Width           =   5685
      Begin VB.OptionButton optFindType 
         Caption         =   "����"
         Height          =   300
         Index           =   4
         Left            =   1500
         TabIndex        =   65
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboPartGroup 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
         Width           =   1980
      End
      Begin VB.ComboBox cboAgeType 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":000C
         Left            =   4545
         List            =   "frmPACSFilter.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1920
         Width           =   825
      End
      Begin VB.TextBox txtEndAge 
         Height          =   330
         Left            =   3855
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtBeginAge 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2400
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox cboSex 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0038
         Left            =   800
         List            =   "frmPACSFilter.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1920
         Width           =   1065
      End
      Begin VB.ComboBox cboAgeWhere 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":003C
         Left            =   3090
         List            =   "frmPACSFilter.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1920
         Width           =   705
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   7
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
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
         Height          =   375
         Index           =   6
         Left            =   4425
         TabIndex        =   12
         Top             =   1320
         Width           =   510
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   5
         Left            =   3945
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "һ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3450
         TabIndex        =   10
         Top             =   1320
         Width           =   510
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   3
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "һ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2475
         TabIndex        =   8
         Top             =   1320
         Width           =   510
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   1
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00C0FFC0&
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboYinYangXing 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":006E
         Left            =   3885
         List            =   "frmPACSFilter.frx":007B
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3975
         Width           =   1530
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "��ͼ"
         Height          =   300
         Index           =   3
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboDiagDOC 
         Height          =   330
         Left            =   3885
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3615
         Width           =   1530
      End
      Begin VB.ComboBox cbo��鼼ʦ 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3135
         Width           =   1530
      End
      Begin VB.ComboBox cboAuditing 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3975
         Width           =   1530
      End
      Begin VB.ComboBox cboModality 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0095
         Left            =   1275
         List            =   "frmPACSFilter.frx":0097
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2400
         Width           =   1980
      End
      Begin VB.ComboBox cbo���� 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0099
         Left            =   3885
         List            =   "frmPACSFilter.frx":00A6
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3135
         Width           =   1530
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "����"
         Height          =   300
         Index           =   2
         Left            =   2230
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.ComboBox cboPart 
         Height          =   330
         Left            =   3360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2760
         Width           =   2055
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "����"
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   840
      End
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3600
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   104726531
         CurrentDate     =   38082.9993055556
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   104726531
         CurrentDate     =   38082
      End
      Begin MSComctlLib.Slider sldDays 
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   529
         _Version        =   393216
         Min             =   1
         Max             =   180
         SelStart        =   1
         TickFrequency   =   7
         Value           =   1
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   5400
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Label Label8 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1650
         TabIndex        =   64
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "��(                             )ʱ�����"
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   270
         Width           =   5055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   210
         Left            =   315
         TabIndex        =   58
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   210
         Left            =   1920
         TabIndex        =   57
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label labYinYangXing 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         Height          =   210
         Left            =   2955
         TabIndex        =   56
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label labDiagDOC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ҽ��"
         Height          =   210
         Left            =   2955
         TabIndex        =   54
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label lab��鼼ʦ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��鼼ʦ"
         Height          =   210
         Left            =   315
         TabIndex        =   53
         Top             =   3195
         Width           =   840
      End
      Begin VB.Label labAuditing 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ҽ��"
         Height          =   210
         Left            =   315
         TabIndex        =   45
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label labModality 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   210
         Left            =   315
         TabIndex        =   44
         Top             =   2460
         Width           =   840
      End
      Begin VB.Label lab���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӱ������"
         Height          =   210
         Left            =   2955
         TabIndex        =   43
         Top             =   3195
         Width           =   840
      End
      Begin VB.Label labPartGroup 
         AutoSize        =   -1  'True
         Caption         =   "��鲿λ"
         Height          =   210
         Left            =   315
         TabIndex        =   41
         Top             =   2820
         Width           =   840
      End
      Begin VB.Label labDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˿���"
         Height          =   210
         Left            =   315
         TabIndex        =   40
         Top             =   3660
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   38
         Top             =   630
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmPACSFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngModul As Long      'ģ��ŵ���
Public mBeforeDays As Integer '��ѯ����
Public mDept As Long
Public mblnOK As Boolean 'ȷ���˳�
Private mrsStudyPart As ADODB.Recordset
Private mrsPartGroup As ADODB.Recordset


Private Sub LoadSex()
'********************************************
'
'��ȡ�����Ա�
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select ���� From �Ա�"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����Ա�")
    CboSex.Clear
    CboSex.AddItem "ȫ��"
        
    With Me.CboSex
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("����"))) & "-" & Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
    End With
    
    CboSex.ListIndex = 0
 
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cboAgeWhere_Click()
    txtBeginAge.Enabled = IIf(cboAgeWhere.Text = "~ ", True, False)
    If txtBeginAge.Enabled Then
        txtBeginAge.BackColor = &HFFFFFF
    Else
        txtBeginAge.Text = ""
        txtBeginAge.BackColor = &HE0E0E0
    End If
End Sub

Private Sub cboModality_Click()
On Error GoTo ErrHandle
    Call FilterGroupPart(cboModality.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cboPartGroup_Click()
On Error GoTo ErrHandle
    Call FilterStudyPart(cboPartGroup.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub chkmore_Click()
    If chkmore.value = 1 Then
        Me.Height = Picmore.Top + Picmore.Height + PicButton.Height + 400
        Picmore.Visible = True
    Else
        Me.Height = Frabase.Top + Frabase.Height + PicButton.Height + 400
        Picmore.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub

Private Sub cmdDayCfg_Click(Index As Integer)
    dtpEnd.value = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59") 'zlDatabase.Currentdate
    
    Select Case Index
        Case 0
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd"))
            sldDays.value = 1
        Case 1
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 1
            sldDays.value = 2
        Case 2
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 6
            sldDays.value = 7
        Case 3
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 13
            sldDays.value = 14
        Case 4
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 29
            sldDays.value = 30
        Case 5
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 59
            sldDays.value = 60
        Case 6
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 89
            sldDays.value = 90
        Case 7
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 179
            sldDays.value = 180
    End Select
End Sub

Private Sub cmdDefault_Click()
    Call Form_Load
    Call Form_Activate
End Sub



Private Sub cmdOK_Click()

    If dtpEnd.value < dtpBegin.value Then
        MsgBoxD Me, "��ʼʱ�䲻�ܴ��ڽ�ֹʱ�䣬���飡", vbInformation, gstrSysName
        dtpEnd.SetFocus
        Exit Sub
    End If
    
    mblnOK = True
    Me.Hide
End Sub

Private Sub GetQuerySchemaCfg(ByRef strSchemaFormat As String, ByRef strSchemaField As String)
    Dim strSchema As String
    Dim intBeginAge As Integer
    Dim intEndAge As Integer
    
    If Not dtpBegin Is Nothing Then
        If optFindType(1).value Then strSchema = strSchema & GetFieldFormatStr("����ʱ��", "����ҽ������", ">=", dtpBegin.value, "And", "")
        If optFindType(2).value Then strSchema = strSchema & GetFieldFormatStr("�״�ʱ��", "����ҽ������", ">=", dtpBegin.value, "And", "")
        If optFindType(3).value Then strSchema = strSchema & GetFieldFormatStr("��������", "Ӱ�����¼", ">=", dtpBegin.value, "And", "")
    End If
        
    If Not dtpEnd Is Nothing Then
        If optFindType(1).value Then strSchema = strSchema & GetFieldFormatStr("����ʱ��", "����ҽ������", "<=", dtpEnd.value, "And", "")
        If optFindType(2).value Then strSchema = strSchema & GetFieldFormatStr("�״�ʱ��", "����ҽ������", "<=", dtpEnd.value, "And", "")
        If optFindType(3).value Then strSchema = strSchema & GetFieldFormatStr("��������", "Ӱ�����¼", "<=", dtpEnd.value, "And", "")
    End If
    
    If CboSex.ListIndex <> 0 Then strSchema = strSchema & GetFieldFormatStr("�Ա�", "������Ϣ", "<=", NeedName(CboSex.Text), "And", "")
    
    
    Select Case NeedName(cboAgeType.Text)
        Case "��"
            intBeginAge = Val(txtBeginAge.Text) * 365
            intEndAge = Val(txtEndAge.Text) * 365
        Case "��"
            intBeginAge = Val(txtBeginAge.Text) * 30
            intEndAge = Val(txtEndAge.Text) * 30
        Case "��"
            intBeginAge = Val(txtBeginAge.Text) * 7
            intEndAge = Val(txtEndAge.Text) * 7
        Case "��"
            intBeginAge = Val(txtBeginAge.Text) * 1
            intEndAge = Val(txtEndAge.Text) * 1
    End Select
        
    If Trim(txtBeginAge.Text) <> "" Then
        If Trim(cboAgeWhere.Text) = "~" Then
            strSchema = strSchema & GetFieldFormatStr("����", "������Ϣ", ">=", CStr(intBeginAge), "And", "")
        End If
    End If
    
    If Trim(txtEndAge.Text) <> "" Then
        If Trim(cboAgeWhere.Text) = "~" Then
            strSchema = strSchema & GetFieldFormatStr("����", "������Ϣ", "<=", CStr(intEndAge), "And", "")
        Else
            strSchema = strSchema & GetFieldFormatStr("����", "������Ϣ", Trim(cboAgeWhere.Text), CStr(intEndAge), "And", "")
        End If
    End If
    
    If cboDept.ListIndex <> 0 Then
        strSchema = strSchema & GetFieldFormatStr("���˿���ID+0", "����ҽ����¼", "=", cboDept.ItemData(cboDept.ListIndex), "And", "")
    End If
    
    If cboPart.ListIndex <> 0 Then
        strSchema = strSchema & GetFieldFormatStr("���˿���ID+0", "����ҽ����¼", "=", cboDept.ItemData(cboDept.ListIndex), "And", "")
    End If
        
End Sub

Private Function GetFieldFormatStr(strFieldName As String, strTabName As String, _
    strWhere As String, strData As String, strLink As String, strQueryType As String, Optional strBracket As String) As String
    Dim strResult As String
            
    On Error GoTo ErrHandle
        strResult = "<" & strFieldName & ">#B=" & strBracket & "#F=" & strTabName & "#W=" & strWhere & "#D=" & strData & "#L=" & strLink & "#T=" & strQueryType & "</" & strFieldName & ">"
        
        GetFieldFormatStr = strResult & vbNewLine
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub GetQueryTimeType()
    'If optFindType(1).value Then GetQueryTimeType = "����ʱ��"
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    'If KeyAscii = vbKeyF2 Then cmdOK_Click
End Sub
Private Sub Form_Activate()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    TxtӰ�����.Text = ""
    txt��������.Text = ""
    'cboYinYangXing.ListIndex = 0
    'cbo����.ListIndex = 0
    'cboSex.ListIndex = 0
    'cboAgeWhere.ListIndex = 4
    'cboAgeType.ListIndex = 0
    txt���.Text = ""
    
    
    '�жϺ���ȡPACS�������ݵ���������
    If pReport_CheckViewName = "" Or pReport_ResultName = "" Or pReport_AdviceName = "" Then
        
        pReport_CheckViewName = "�������"
        pReport_ResultName = "������"
        pReport_AdviceName = "����"
        
        strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1] " _
            & " and (������ = '�����������' or ������ = '����������' or ������ = '��������') "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mDept)
        
        If rsTemp.EOF = False Then
            Select Case rsTemp!������
                Case "�����������"
                    pReport_CheckViewName = Nvl(rsTemp!����ֵ, "�������")
                Case "����������"
                    pReport_ResultName = Nvl(rsTemp!����ֵ, "������")
                Case "��������"
                    pReport_AdviceName = Nvl(rsTemp!����ֵ, "����")
            End Select
        End If
    End If
    txtPacsRpt(0).Text = ""
    txtPacsRpt(1).Text = ""
    txtPacsRpt(2).Text = ""
    lblPacsRpt(0).Caption = pReport_CheckViewName
    lblPacsRpt(1).Caption = "�� " & pReport_ResultName
    lblPacsRpt(2).Caption = "�� " & pReport_AdviceName
        
    dtpBegin.SetFocus
End Sub
Private Sub LoadDept()
'���ܣ����ݲ�����Դ��ȡ���˿���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select Distinct A.ID,A.����,A.����,B.�������" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.�������� IN('�ٴ�','����')" & _
        " And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cboDept.Clear
    cboDept.AddItem "���п���"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitModality()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select ����,���� from ���Ƽ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ƽ������")
    
    cboModality.Clear
    cboModality.AddItem "�������"
    cboModality.ListIndex = 0
    
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!���� & "-" & rsTemp!����
        rsTemp.MoveNext
    Loop
    
End Sub

Private Function InitPart() As Boolean
    Dim strSQL As String
    Dim strGroup As String
    Dim strStudyPart As String
    Dim strRead As String
    Dim strGroupPY  As String
    Dim strFilter As String
    
    strFilter = ""
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        strFilter = " where ����='����' or upper(����)='DG' or upper(����)='BL'"
    End If
    
    '��ȡ��λ����
    strSQL = "select Distinct ���� || '(' || ���� || ')' as ����, ���� from ���Ƽ�鲿λ " & strFilter & " order by ����"
    Set mrsPartGroup = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��λ")
    
    cboPartGroup.Clear
    cboPartGroup.AddItem "���з���"
    cboPartGroup.ListIndex = 0
    
    While Not mrsPartGroup.EOF
        strGroup = Nvl(mrsPartGroup("����"))
        
         '��ȡ�������� ��ת��Ϊƴ������
        strGroupPY = zlCommFun.SpellCode(Mid(strGroup, InStr(strGroup, "-") + 1, InStrRev(strGroup, "(") - InStr(strGroup, "-") - 1))
        
        strGroup = IIf(InStr(strGroup, "-") = 0, "-" & strGroup, strGroup)
        
        strGroup = strGroupPY & Mid(strGroup, InStr(strGroup, "-"), Len(strGroup))
        
        cboPartGroup.AddItem strGroup
        mrsPartGroup.MoveNext
    Wend
    

    
    '��ȡ��鲿λ
    strSQL = "Select Distinct ����, ���� || '(' || ���� || ')'  as ����, ���� From ���Ƽ�鲿λ " & strFilter & " order by ����"
    Set mrsStudyPart = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��λ")
    
    cboPart.Clear
    cboPart.AddItem "���в�λ"
    cboPart.ListIndex = 0
    
    strRead = ""
    
    With Me.cboPart
        Do While Not mrsStudyPart.EOF
            strStudyPart = zlCommFun.SpellCode(Nvl(mrsStudyPart("����"))) & "-" & Nvl(mrsStudyPart("����"))
            
            If InStr(strRead, strStudyPart & ";") <= 0 Then
                .AddItem strStudyPart
                
                strRead = strRead & strStudyPart & ";"
            End If
            
            mrsStudyPart.MoveNext
        Loop
    End With
    

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub FilterStudyPart(ByVal strGroupName As String)
    Dim strStudyPart As String
    Dim strRead As String
    Dim strSQL As String
    Dim strGroup As String
    Dim strType As String
    Dim rsGroup As ADODB.Recordset
    
     If mrsStudyPart Is Nothing Then Exit Sub
    
    If strGroupName <> "���з���" Then
        '�����ʱ������ǰ׺�ı�� �滻����ƴ�����룬����Ҫ
        strGroup = Mid(strGroupName, InStr(strGroupName, "-") + 1, InStrRev(strGroupName, "(") - InStr(strGroupName, "-") - 1)
        strType = Mid(strGroupName, InStr(strGroupName, "(") + 1, InStrRev(strGroupName, ")") - InStr(strGroupName, "(") - 1)
    
        strSQL = "Select Distinct  ���� From ���Ƽ�鲿λ where ���� = '" & strType & "' and ���� like '%" & strGroup & "%'"
        Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, "�õ�������Ϣ")

        strGroup = IIf(rsGroup.RecordCount < 1, "", Nvl(rsGroup!����) & "(" & strType & ")")
    End If
    
    mrsStudyPart.Filter = IIf(strGroupName = "���з���", "", "����='" & strGroup & "'")
    
    cboPart.Clear
    cboPart.AddItem "���в�λ"
    cboPart.ListIndex = 0
    
    With Me.cboPart
        Do While Not mrsStudyPart.EOF
            strStudyPart = zlCommFun.SpellCode(Nvl(mrsStudyPart("����"))) & "-" & Nvl(mrsStudyPart("����"))
            
            If InStr(strRead, strStudyPart & ";") <= 0 Then
                .AddItem strStudyPart
                
                strRead = strRead & strStudyPart & ";"
            End If
            
            mrsStudyPart.MoveNext
        Loop
    End With

End Sub

Private Sub FilterGroupPart(ByVal strTypeName As String)
'���˲�λ����
    Dim strGroupPart As String
    Dim strRead As String
    Dim strType As String
    Dim strGroupPY As String
    
    If mrsPartGroup Is Nothing Then Exit Sub
    
    strType = Mid(strTypeName, InStr(strTypeName, "-") + 1, Len(strTypeName))

    mrsPartGroup.Filter = IIf(strTypeName = "�������", "", "����='" & strType & "'")
    
    cboPartGroup.Clear
    cboPartGroup.AddItem "���з���"
    cboPartGroup.ListIndex = 0
    
    With Me.cboPartGroup
        Do While Not mrsPartGroup.EOF
            strGroupPart = Nvl(mrsPartGroup("����"))
            
            '��ȡ�������� ��ת��Ϊƴ������
            strGroupPY = zlCommFun.SpellCode(Mid(strGroupPart, InStr(strGroupPart, "-") + 1, InStrRev(strGroupPart, "(") - InStr(strGroupPart, "-") - 1))
            
            strGroupPart = IIf(InStr(strGroupPart, "-") = 0, "-" & strGroupPart, strGroupPart)
            
            strGroupPart = strGroupPY & Mid(strGroupPart, InStr(strGroupPart, "-"), Len(strGroupPart))
        
            .AddItem strGroupPart
            
            mrsPartGroup.MoveNext
        Loop
    End With
    
    
End Sub




Private Sub InitDoc()
Dim rsTmp As ADODB.Recordset
    cboDiagDOC.Clear: cboAuditing.Clear: cbo��鼼ʦ.Clear
    cboDiagDOC.AddItem "����ҽ��": cboAuditing.AddItem "����ҽ��": cbo��鼼ʦ.AddItem "����ҽ��"
    cboDiagDOC.ListIndex = 0: cboAuditing.ListIndex = 0: cbo��鼼ʦ.ListIndex = 0
    On Error GoTo errH
    gstrSQL = "select distinct A.����,A.���� from ��Ա�� A,������Ա B where B.����ID=[1] AND A.ID=B.��ԱID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҽ��", mDept)
    If rsTmp Is Nothing Then Exit Sub
    Do While Not rsTmp.EOF
        cboDiagDOC.AddItem rsTmp!���� & "-" & rsTmp!����
        cboAuditing.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo��鼼ʦ.AddItem rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optAdviceTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub optCheckTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub Form_Load()
Dim curDate As Date
Dim intʱ������ As Integer

    curDate = zlDatabase.Currentdate
    dtpEnd.value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpEnd.Tag = dtpEnd.value
    dtpBegin.value = Format(dtpEnd.value - mBeforeDays, "yyyy-MM-dd 00:00")
    
    intʱ������ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", 1))
    If intʱ������ = 1 Then
        optFindType(1).value = True
    ElseIf intʱ������ = 2 Then
        optFindType(2).value = True
    ElseIf intʱ������ = 3 Then
        optFindType(3).value = True
    Else
        optFindType(4).value = True
    End If
    
    '����λ��ʱ������Me.Visible���������ô���ֻ�ڵ�һ�μ���ʱ���Ž��е���
    '����
    If mlngModul = G_LNG_PATHOLSYS_NUM And Not Me.Visible Then
        optFindType(1).Caption = "����"
        optFindType(3).Caption = "��������"
        
        
        labModality.Visible = False
        cboModality.Visible = False
        
        lab��鼼ʦ.Visible = False
        cbo��鼼ʦ.Visible = False
        
        lab����.Visible = False
        cbo����.Visible = False
        
        labPartGroup.Top = labPartGroup.Top - cboModality.Height
        cboPartGroup.Top = cboPartGroup.Top - cboModality.Height
        cboPart.Top = cboPart.Top - cboModality.Height
        
        labDept.Top = labDept.Top - cbo��鼼ʦ.Height - cboModality.Height
        cboDept.Top = cboDept.Top - cbo��鼼ʦ.Height - cboModality.Height
        
        labDiagDOC.Top = labDiagDOC.Top - cbo��鼼ʦ.Height - cboModality.Height
        cboDiagDOC.Top = cboDiagDOC.Top - cbo��鼼ʦ.Height - cboModality.Height
        
        labAuditing.Top = labAuditing.Top - cbo��鼼ʦ.Height - cboModality.Height
        cboAuditing.Top = cboAuditing.Top - cbo��鼼ʦ.Height - cboModality.Height
        
        labYinYangXing.Top = labYinYangXing.Top - cbo��鼼ʦ.Height - cboModality.Height
        cboYinYangXing.Top = cboYinYangXing.Top - cbo��鼼ʦ.Height - cboModality.Height
        
        Frabase.Height = Frabase.Height - cbo��鼼ʦ.Height - cboModality.Height
        Picmore.Top = Picmore.Top - cbo��鼼ʦ.Height - cboModality.Height
        
        Me.Height = Me.Height - cbo��鼼ʦ.Height - cboModality.Height
        
    '�ɼ�
    ElseIf mlngModul = G_LNG_VIDEOSTATION_MODULE And Not Me.Visible Then
        labModality.Visible = False
        cboModality.Visible = False
        
        labPartGroup.Top = labPartGroup.Top - cboModality.Height
        cboPartGroup.Top = cboPartGroup.Top - cboModality.Height
        cboPart.Top = cboPart.Top - cboModality.Height
        
        labDept.Top = labDept.Top - cboModality.Height
        cboDept.Top = cboDept.Top - cboModality.Height
        
        labDiagDOC.Top = labDiagDOC.Top - cboModality.Height
        cboDiagDOC.Top = cboDiagDOC.Top - cboModality.Height
        
        labAuditing.Top = labAuditing.Top - cboModality.Height
        cboAuditing.Top = cboAuditing.Top - cboModality.Height
        
        labYinYangXing.Top = labYinYangXing.Top - cboModality.Height
        cboYinYangXing.Top = cboYinYangXing.Top - cboModality.Height
        
        Frabase.Height = Frabase.Height - cboModality.Height
        Picmore.Top = Picmore.Top - cboModality.Height
        
        lab��鼼ʦ.Top = lab��鼼ʦ.Top - cboModality.Height
        cbo��鼼ʦ.Top = cbo��鼼ʦ.Top - cboModality.Height
        
        lab����.Top = lab����.Top - cboModality.Height
        cbo����.Top = cbo����.Top - cboModality.Height
        
        Me.Height = Me.Height - cboModality.Height
    'ҽ��
    Else
        '......
    End If
    
    'ֻ�ڴ����һ�μ���ʱִ�У�������ȱʡ���ܰ�ť�����ظ����м���
    If Not Me.Visible Then
        LoadDept
        LoadSex
        InitModality
        InitPart
        InitDoc
    End If
    
    
    txtBeginAge.Text = ""
    txtEndAge.Text = ""
    
    cboModality.Text = "�������"
    cboPartGroup.Text = "���з���"
    cboPart.Text = "���в�λ"
    cboDept.Text = "���п���"
    cboDiagDOC.Text = "����ҽ��"
    cboAuditing.Text = "����ҽ��"
    cbo��鼼ʦ.Text = "����ҽ��"
    
    cboYinYangXing.ListIndex = 0
    cbo����.ListIndex = 0
    CboSex.ListIndex = 0
    cboAgeWhere.ListIndex = 0
    cboAgeType.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim intʱ������ As Integer
    '�������ʱ������
    If optFindType(1).value = True Then
        intʱ������ = 1
    ElseIf optFindType(2).value = True Then
        intʱ������ = 2
    ElseIf optFindType(3).value = True Then
        intʱ������ = 3
    Else
        intʱ������ = 4
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", intʱ������
End Sub



Private Sub sldDays_Scroll()
    '���ò�ѯʱ�䷶Χ
    dtpBegin.value = Format(CDate(dtpEnd.value - (sldDays.value - 1)), "yyyy-mm-dd 00:00")
End Sub
