VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
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
   ScaleHeight     =   5355
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6345
      TabIndex        =   22
      Top             =   4740
      Width           =   6345
      Begin VB.ComboBox cboSchemaName 
         Height          =   330
         Left            =   1080
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelSchema 
         Caption         =   "ɾ������(&D)"
         Height          =   375
         Left            =   3480
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveSchema 
         Caption         =   "���淽��(&S)"
         Height          =   375
         Left            =   4920
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkmore 
         Caption         =   "��������(&M)"
         Height          =   375
         Left            =   1440
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1305
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "ȱʡ(&F)"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   375
         Left            =   5025
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   375
         Left            =   3660
         TabIndex        =   14
         ToolTipText     =   "F2"
         Top             =   120
         Width           =   1215
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
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   6345
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   6345
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
         TabIndex        =   38
         Top             =   405
         Width           =   4410
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
         TabIndex        =   10
         Top             =   30
         Width           =   4410
      End
      Begin VB.Frame Frame1 
         Caption         =   "PACS�����ѯ"
         Height          =   1455
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   5850
         Begin VB.TextBox txtPacsRpt 
            Height          =   300
            Index           =   0
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   4215
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   16
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   300
            Index           =   2
            Left            =   1440
            TabIndex        =   13
            Top             =   960
            Width           =   4215
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "�������"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "�� ������"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "�� ����"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   30
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
         TabIndex        =   11
         Top             =   795
         Width           =   4410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   465
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   210
         Left            =   240
         TabIndex        =   34
         Top             =   90
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   210
         Left            =   240
         TabIndex        =   33
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
      Height          =   4860
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   7005
      Begin VB.ComboBox cboAgeType 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":000C
         Left            =   4800
         List            =   "frmPACSFilter.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2640
         Width           =   1425
      End
      Begin VB.TextBox txtEndAge 
         Height          =   330
         Left            =   3840
         TabIndex        =   57
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtBeginAge 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1275
         TabIndex        =   56
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cboSex 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0038
         Left            =   1275
         List            =   "frmPACSFilter.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2280
         Width           =   1665
      End
      Begin VB.ComboBox cboAgeWhere 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":003C
         Left            =   2280
         List            =   "frmPACSFilter.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   2640
         Width           =   1425
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "������(&Y)"
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   51
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "������(&H)"
         Height          =   375
         Index           =   6
         Left            =   3480
         TabIndex        =   50
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "������(&U)"
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   49
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "��һ��(&N)"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   48
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "������(&K)"
         Height          =   375
         Index           =   3
         Left            =   5040
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "��һ��(&W)"
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   46
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "������(&A)"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   45
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "����(&T)"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cboYinYangXing 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":006E
         Left            =   4365
         List            =   "frmPACSFilter.frx":007B
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   4320
         Width           =   1890
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "����ͼʱ�����"
         Height          =   300
         Index           =   3
         Left            =   4560
         TabIndex        =   37
         Top             =   240
         Width           =   1800
      End
      Begin VB.ComboBox cbodiagdoc 
         Height          =   330
         Left            =   4365
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3960
         Width           =   1890
      End
      Begin VB.ComboBox cbo��鼼ʦ 
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3975
         Width           =   1905
      End
      Begin VB.ComboBox cboAuditing 
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4350
         Width           =   1905
      End
      Begin VB.ComboBox cboModality 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0095
         Left            =   4365
         List            =   "frmPACSFilter.frx":0097
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3615
         Width           =   1890
      End
      Begin VB.ComboBox cbo���� 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0099
         Left            =   1275
         List            =   "frmPACSFilter.frx":00A6
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3615
         Width           =   1905
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "������ʱ�����"
         Height          =   300
         Index           =   2
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.ComboBox cboPart 
         Height          =   330
         Left            =   4365
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3240
         Width           =   1890
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "������ʱ�����"
         Height          =   300
         Index           =   1
         Left            =   315
         TabIndex        =   18
         Top             =   240
         Width           =   1800
      End
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   4080
         TabIndex        =   2
         Top             =   570
         Width           =   2190
         _ExtentX        =   3863
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
         Format          =   89260035
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2070
         _ExtentX        =   3651
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
         Format          =   89260035
         CurrentDate     =   38082
      End
      Begin MSComctlLib.Slider sldDays 
         Height          =   300
         Left            =   1320
         TabIndex        =   43
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   529
         _Version        =   393216
         Max             =   180
         TickFrequency   =   7
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա�"
         Height          =   210
         Left            =   315
         TabIndex        =   55
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   210
         Left            =   315
         TabIndex        =   54
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label 
         Caption         =   "ʱ�䷶Χ"
         Height          =   255
         Index           =   0
         Left            =   315
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� �� ��"
         Height          =   210
         Left            =   3435
         TabIndex        =   40
         Top             =   4440
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ҽ��"
         Height          =   210
         Left            =   3435
         TabIndex        =   36
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��鼼ʦ"
         Height          =   210
         Left            =   315
         TabIndex        =   35
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ҽ��"
         Height          =   210
         Left            =   315
         TabIndex        =   27
         Top             =   4410
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӱ�����"
         Height          =   210
         Left            =   3435
         TabIndex        =   26
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӱ������"
         Height          =   210
         Left            =   315
         TabIndex        =   25
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "��鲿λ"
         Height          =   210
         Left            =   3435
         TabIndex        =   21
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˿���"
         Height          =   210
         Left            =   315
         TabIndex        =   20
         Top             =   3300
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3720
         TabIndex        =   19
         Top             =   600
         Width           =   180
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


Private Sub LoadSex()
'********************************************
'
'��ȡ�����Ա�
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "Select ���� From �Ա�"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����Ա�")
    cboSex.Clear
    cboSex.AddItem "ȫ��"
        
    With Me.cboSex
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("����"))) & "-" & Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
    End With
    
    cboSex.ListIndex = 0
 
    Exit Sub
    
errHandle:
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

Private Sub chkmore_Click()
    If chkmore.Value = 1 Then
        Me.Height = Picmore.Top + Picmore.Height + PicButton.Height + 500
        Picmore.Visible = True
    Else
        Me.Height = Frabase.Top + Frabase.Height + PicButton.Height + 500
        Picmore.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub

Private Sub cmdDayCfg_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpBegin.Value = CDate(Now - Time)
            sldDays.Value = 0
        Case 1
            dtpBegin.Value = CDate(dtpEnd.Value - 2)
            sldDays.Value = 2
        Case 2
            dtpBegin.Value = CDate(dtpEnd.Value - 7)
            sldDays.Value = 7
        Case 3
            dtpBegin.Value = CDate(dtpEnd.Value - 14)
            sldDays.Value = 14
        Case 4
            dtpBegin.Value = CDate(dtpEnd.Value - 30)
            sldDays.Value = 30
        Case 5
            dtpBegin.Value = CDate(dtpEnd.Value - 60)
            sldDays.Value = 60
        Case 6
            dtpBegin.Value = CDate(dtpEnd.Value - 90)
            sldDays.Value = 90
        Case 7
            dtpBegin.Value = CDate(dtpEnd.Value - 180)
            sldDays.Value = 180
    End Select
End Sub

Private Sub cmdDefault_Click()
    Call Form_Load
    Call Form_Activate
End Sub

Private Sub cmdOK_Click()

    If dtpEnd.Value < dtpBegin.Value Then
        MsgBox "��ʼʱ�䲻�ܴ��ڽ�ֹʱ�䣬���飡", vbInformation, gstrSysName
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
        If optFindType(1).Value Then strSchema = strSchema & GetFieldFormatStr("����ʱ��", "����ҽ������", ">=", dtpBegin.Value, "And", "")
        If optFindType(2).Value Then strSchema = strSchema & GetFieldFormatStr("�״�ʱ��", "����ҽ������", ">=", dtpBegin.Value, "And", "")
        If optFindType(3).Value Then strSchema = strSchema & GetFieldFormatStr("��������", "Ӱ�����¼", ">=", dtpBegin.Value, "And", "")
    End If
        
    If Not dtpEnd Is Nothing Then
        If optFindType(1).Value Then strSchema = strSchema & GetFieldFormatStr("����ʱ��", "����ҽ������", "<=", dtpEnd.Value, "And", "")
        If optFindType(2).Value Then strSchema = strSchema & GetFieldFormatStr("�״�ʱ��", "����ҽ������", "<=", dtpEnd.Value, "And", "")
        If optFindType(3).Value Then strSchema = strSchema & GetFieldFormatStr("��������", "Ӱ�����¼", "<=", dtpEnd.Value, "And", "")
    End If
    
    If cboSex.ListIndex <> 0 Then strSchema = strSchema & GetFieldFormatStr("�Ա�", "������Ϣ", "<=", NeedName(cboSex.Text), "And", "")
    
    
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
            
    On Error GoTo errHandle
        strResult = "<" & strFieldName & ">#B=" & strBracket & "#F=" & strTabName & "#W=" & strWhere & "#D=" & strData & "#L=" & strLink & "#T=" & strQueryType & "</" & strFieldName & ">"
        
        GetFieldFormatStr = strResult & vbNewLine
    Exit Function
errHandle:
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
    If KeyAscii = vbKeyF2 Then cmdOK_Click
End Sub
Private Sub Form_Activate()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim pReport_CheckViewName As String
    Dim pReport_ResultName As String
    Dim pReport_AdviceName As String
    
    TxtӰ�����.Text = ""
    txt��������.Text = ""
    'cboYinYangXing.ListIndex = 0
    'cbo����.ListIndex = 0
    'cboSex.ListIndex = 0
    'cboAgeWhere.ListIndex = 4
    'cboAgeType.ListIndex = 0
    txt���.Text = ""
    
    
        
        pReport_CheckViewName = "�������"
        pReport_ResultName = "������"
        pReport_AdviceName = "����"

    txtPacsRpt(0).Text = ""
    txtPacsRpt(1).Text = ""
    txtPacsRpt(2).Text = ""
    lblPacsRpt(0).Caption = pReport_CheckViewName
    lblPacsRpt(1).Caption = "�� " & pReport_ResultName
    lblPacsRpt(2).Caption = "�� " & pReport_AdviceName
    
    If mlngModul = 1290 Then    'Ӱ��ҽ������վ
        Label15.Visible = True
        cboModality.Visible = True
    ElseIf mlngModul = 1291 Then    'Ӱ��ɼ�����վ
        Label15.Visible = False
        cboModality.Visible = False
    ElseIf mlngModul = 1293 Then        'Ӱ������վ
        Label15.Visible = False
        cboModality.Visible = False
    End If
        
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

Private Function InitPart() As Boolean
    Dim rsTmp As ADODB.Recordset
    gstrSQL = "Select Distinct ���� From ���Ƽ�鲿λ"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ")
    cboPart.Clear
    cboPart.AddItem "���в�λ"
    cboPart.ListIndex = 0
    With Me.cboPart
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("����"))) & "-" & Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
    End With
 
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub InitDoc()
Dim rsTmp As ADODB.Recordset
    cbodiagdoc.Clear: cboAuditing.Clear: cbo��鼼ʦ.Clear
    cbodiagdoc.AddItem "����ҽ��": cboAuditing.AddItem "����ҽ��": cbo��鼼ʦ.AddItem "����ҽ��"
    cbodiagdoc.ListIndex = 0: cboAuditing.ListIndex = 0: cbo��鼼ʦ.ListIndex = 0
    On Error GoTo errH
    gstrSQL = "select distinct A.����,A.���� from ��Ա�� A,������Ա B where B.����ID=[1] AND A.ID=B.��ԱID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҽ��", mDept)
    If rsTmp Is Nothing Then Exit Sub
    Do While Not rsTmp.EOF
        cbodiagdoc.AddItem rsTmp!���� & "-" & rsTmp!����
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
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpEnd.Tag = dtpEnd.Value
    dtpBegin.Value = Format(dtpEnd.Value - mBeforeDays, "yyyy-MM-dd 00:00")
    
    intʱ������ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", 1))
    If intʱ������ = 1 Then
        optFindType(1).Value = True
    ElseIf intʱ������ = 2 Then
        optFindType(2).Value = True
    Else
        optFindType(3).Value = True
    End If
    
    LoadDept
    LoadSex
    InitPart
    InitDoc
    InitModality
    
    cboYinYangXing.ListIndex = 0
    cbo����.ListIndex = 0
    cboSex.ListIndex = 0
    cboAgeWhere.ListIndex = 0
    cboAgeType.ListIndex = 0
End Sub
Private Sub InitModality()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select ����,���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Ӱ�������")
    
    cboModality.Clear
    cboModality.AddItem "ȫ��"
    
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!���� & "--" & rsTemp!����
        rsTemp.MoveNext
    Loop
    cboModality.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intʱ������ As Integer
    '�������ʱ������
    If optFindType(1).Value = True Then
        intʱ������ = 1
    ElseIf optFindType(2).Value = True Then
        intʱ������ = 2
    Else
        intʱ������ = 3
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", intʱ������
End Sub

Private Sub sldDays_Scroll()
    '���ò�ѯʱ�䷶Χ
    dtpBegin.Value = CDate(dtpEnd.Value - sldDays.Value)
End Sub
