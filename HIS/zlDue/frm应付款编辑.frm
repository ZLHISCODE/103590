VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmӦ����༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ����༭"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frmӦ����༭.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      Height          =   5250
      Left            =   30
      ScaleHeight     =   5190
      ScaleWidth      =   9420
      TabIndex        =   49
      Top             =   60
      Width           =   9480
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   915
         TabIndex        =   1
         Top             =   735
         Width           =   4200
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   13
         Left            =   975
         MaxLength       =   50
         TabIndex        =   38
         Tag             =   "ժҪ"
         Top             =   4095
         Width           =   8340
      End
      Begin VB.CommandButton cmdSelDept 
         Caption         =   "��"
         Height          =   300
         Left            =   5100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
      End
      Begin VB.Frame fraTemp 
         Height          =   3045
         Left            =   75
         TabIndex        =   50
         Top             =   990
         Width           =   9255
         Begin VB.OptionButton optClass 
            Caption         =   "�豸(&4)"
            Height          =   180
            Index           =   3
            Left            =   6120
            TabIndex        =   7
            Top             =   220
            Width           =   1000
         End
         Begin VB.OptionButton optClass 
            Caption         =   "����(&3)"
            Height          =   180
            Index           =   2
            Left            =   5040
            TabIndex        =   6
            Top             =   220
            Width           =   1000
         End
         Begin VB.OptionButton optClass 
            Caption         =   "����(&2)"
            Height          =   180
            Index           =   1
            Left            =   3960
            TabIndex        =   5
            Top             =   220
            Width           =   1000
         End
         Begin VB.OptionButton optClass 
            Caption         =   "ҩƷ(&1)"
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   4
            Top             =   220
            Width           =   1000
         End
         Begin VB.CheckBox chkSelector 
            Caption         =   "ѡ����¼��(&D)"
            Height          =   180
            Left            =   1110
            TabIndex        =   3
            Top             =   220
            Width           =   1500
         End
         Begin VB.CommandButton cmdSelName 
            Caption         =   "��"
            Height          =   300
            Left            =   8835
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   480
            Width           =   300
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   14
            Left            =   1110
            MaxLength       =   200
            TabIndex        =   18
            Tag             =   "�������"
            Top             =   1560
            Width           =   4875
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   9
            Left            =   7260
            MaxLength       =   16
            TabIndex        =   30
            Tag             =   "��Ʊ���"
            Top             =   2250
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   7260
            MaxLength       =   8
            TabIndex        =   20
            Tag             =   "��ⵥ�ݺ�"
            Top             =   1560
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   7
            Left            =   4110
            MaxLength       =   16
            TabIndex        =   24
            Tag             =   "���ݽ��"
            Top             =   1920
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   9
            Tag             =   "Ʒ��"
            Top             =   480
            Width           =   7710
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "���"
            Top             =   840
            Width           =   8025
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   14
            Tag             =   "����"
            Top             =   1200
            Width           =   4890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   7260
            MaxLength       =   8
            TabIndex        =   16
            Tag             =   "������λ"
            Top             =   1200
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   10
            Left            =   1110
            MaxLength       =   16
            TabIndex        =   32
            Tag             =   "����"
            Top             =   2640
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   12
            Left            =   7260
            MaxLength       =   16
            TabIndex        =   36
            Tag             =   "�ɹ����"
            Top             =   2640
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   8
            Left            =   1110
            MaxLength       =   200
            TabIndex        =   28
            Tag             =   "��Ʊ��"
            Top             =   2280
            Width           =   4875
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   1110
            MaxLength       =   20
            TabIndex        =   22
            Tag             =   "����"
            Top             =   1920
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   4110
            MaxLength       =   16
            TabIndex        =   34
            Tag             =   "�ɹ���"
            Top             =   2640
            Width           =   1875
         End
         Begin MSComCtl2.DTPicker Dtp��Ʊ���� 
            Height          =   300
            Left            =   7260
            TabIndex        =   26
            Tag             =   "��Ʊ����"
            Top             =   1920
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   314638336
            CurrentDate     =   37904
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "�������(&B)"
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   17
            Top             =   1620
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "��ⵥ��(&R)"
            Height          =   180
            Index           =   3
            Left            =   6195
            TabIndex        =   19
            Top             =   1620
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "���ݽ��(&J)"
            Height          =   180
            Index           =   4
            Left            =   3105
            TabIndex        =   23
            Top             =   1960
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ���(&E)"
            Height          =   180
            Index           =   1
            Left            =   6195
            TabIndex        =   29
            Top             =   2310
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ����(&F)"
            Height          =   180
            Index           =   16
            Left            =   6195
            TabIndex        =   25
            Top             =   1960
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Ʒ��(&N)"
            Height          =   180
            Index           =   7
            Left            =   465
            TabIndex        =   8
            Top             =   530
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "���(&G)"
            Height          =   180
            Index           =   8
            Left            =   465
            TabIndex        =   11
            Top             =   890
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "����(&A)"
            Height          =   180
            Index           =   9
            Left            =   465
            TabIndex        =   13
            Top             =   1240
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "������λ(&U)"
            Height          =   180
            Index           =   11
            Left            =   6195
            TabIndex        =   15
            Top             =   1240
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "����(&S)"
            Height          =   180
            Index           =   12
            Left            =   465
            TabIndex        =   31
            Top             =   2700
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "�ɹ����(&I)"
            Height          =   180
            Index           =   14
            Left            =   6195
            TabIndex        =   35
            Top             =   2700
            Width           =   990
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ��(&K)"
            Height          =   180
            Index           =   0
            Left            =   285
            TabIndex        =   27
            Top             =   2340
            Width           =   810
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "����(&P)"
            Height          =   180
            Index           =   10
            Left            =   465
            TabIndex        =   21
            Top             =   1960
            Width           =   630
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "�ɹ���(&T)"
            Height          =   180
            Index           =   13
            Left            =   3300
            TabIndex        =   33
            Tag             =   "�ɹ���"
            Top             =   2700
            Width           =   810
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
         Height          =   3645
         Left            =   705
         TabIndex        =   55
         Top             =   5115
         Visible         =   0   'False
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   6429
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   975
         TabIndex        =   40
         Top             =   4470
         Width           =   1875
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   975
         TabIndex        =   44
         Top             =   4830
         Width           =   1875
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   6645
         TabIndex        =   45
         Top             =   4890
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   375
         TabIndex        =   43
         Top             =   4890
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   6645
         TabIndex        =   41
         Top             =   4530
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   375
         TabIndex        =   39
         Top             =   4530
         Width           =   540
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7425
         TabIndex        =   42
         Top             =   4470
         Width           =   1875
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7425
         TabIndex        =   46
         Top             =   4830
         Width           =   1875
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "ժҪ(&W)"
         Height          =   180
         Index           =   5
         Left            =   315
         TabIndex        =   37
         Top             =   4155
         Width           =   630
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ����¼��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   90
         TabIndex        =   53
         Top             =   15
         Width           =   8850
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   6945
         TabIndex        =   52
         Top             =   390
         Width           =   1140
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8040
         TabIndex        =   51
         Top             =   345
         Width           =   1290
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "��Ӧ��(&M)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   0
         Top             =   810
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8325
      TabIndex        =   48
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7050
      TabIndex        =   47
      Top             =   5520
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   54
      Top             =   5940
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmӦ����༭.frx":030A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12250
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmӦ����༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mEditType As gEditType
Private mint��¼״̬ As Integer        '  RecBillStatus       '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mErrBillStatusInfor As ErrBillStatusInfor       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mblnEdit As Boolean             '�༭״̬
Private mblnSuccess As Boolean          '�Ƿ��е��ݱ���ɹ�
Private mstrPrivs  As String
Private mstrNo As String                   '���ݺ�
Private mlng��λID As Long
Private mint��Ʊ��Len As Integer          '���ݿⳤ��
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSave As Boolean
Private mlngID As Long      '����ID
Private mfrmMain  As Object
Private mstrSelectTag As String
Private mintPreCol As Integer
Private mintsort As Integer
Private Const mlngModule = 1322
'�������������
Private Function GetDepend() As Boolean
    Dim strSQL As String
    Dim rsDepend As New Recordset

    GetDepend = False
    Dim strȨ�� As String
    strȨ�� = " and  " & Get����Ȩ��(mstrPrivs)
    strSQL = "" & _
        "   SELECT  Id " & _
        "   FROM ��Ӧ�� " & _
        "   Where   ĩ��=1 " & zl_��ȡվ������ & "  " & strȨ��
    Err = 0
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsDepend, strSQL, Me.Caption
    If rsDepend.EOF Then
        ShowMsgbox "û�����ù�Ӧ�̻�Ȩ�޲��㣬���ڹ�Ӧ�̹��������ã�"
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowCard(FrmMain As Form, ByVal lngID As Long, _
    ByVal int�༭״̬ As gEditType, ByVal strPrivs As String, Optional lng��λID As Long = 0, _
    Optional int��¼״̬ As Integer = 1, _
    Optional blnSuccess As Boolean = False)
    
    mblnSave = False
    mblnSuccess = False
    mEditType = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mlngID = lngID
    mlng��λID = lng��λID
    mblnSuccess = blnSuccess
    mblnChange = False
    mErrBillStatusInfor = 1
    
    mstrPrivs = strPrivs
    
    Set mfrmMain = FrmMain
    
    '�������������ϵ
    If Not GetDepend Then Exit Sub
    
    If mEditType = g���� Then
        mblnEdit = True
    ElseIf mEditType = g�޸� Then
        mblnEdit = True
    ElseIf mEditType = g��� Then
        mblnEdit = False
        cmdOK.Caption = "���(&V)"
    ElseIf mEditType = gȡ�� Then
        mblnEdit = False
        cmdOK.Caption = "����(&O)"
    ElseIf mEditType = g�鿴 Then
        mblnEdit = False
        cmdOK.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            cmdOK.Visible = False
        Else
            cmdOK.Visible = True
        End If
    End If
     
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim strSQL As String
    Dim rsInitCard As New Recordset
    
    On Error GoTo errHandle
    Select Case mEditType
        Case g����
            InitControl
            Txt������ = gstrUserName
            Txt�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            Txt����� = ""
            Txt������� = ""
            
            If mlng��λID = 0 Then Exit Sub
            'ȷ����Ӧ��
            'by lesfeng 2009-12-2 �����Ż�
            Dim strȨ�� As String
            strȨ�� = " and  " & Get����Ȩ��(mstrPrivs)
            
            strSQL = "Select ID,����,����,���� From ��Ӧ��  where id=[1]" & strȨ��
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��λID)
            
            If rsInitCard.EOF Then Exit Sub
            
            txtEdit(0).Text = "[" & Nvl(rsInitCard!����) & "]" & Nvl(rsInitCard!����)
            chkSelector.Tag = Nvl(rsInitCard!����)
            mlng��λID = Nvl(rsInitCard!ID, 0)
        Case g���, g�޸�, g�鿴, gȡ��
            InitControl
            'by lesfeng 2009-12-2 �����Ż�
            strSQL = "" & _
                  " SELECT A.ID,A.��¼����,A.��¼״̬,A.NO,A.��ĿID,A.���,A.�շ�ID,A.��λID,A.Ʒ��,A.���,A.����,A.����,A.������λ," & _
                  "   A.��ⵥ�ݺ�,A.���ݽ��,A.����,A.�ɹ���,A.�ɹ����,A.��Ʊ��,A.��Ʊ����,A.��Ʊ���,A.�ƶ�����,A.�ƻ����," & _
                  "   A.�ƻ���,A.�ƻ�����,A.������,A.��������,A.�����,A.�������,A.ժҪ,A.�������,A.�ƻ����,A.ϵͳ��ʶ,A.�������," & _
                  "   b.���� as ��Ӧ�̱���,b.���� as ��Ӧ������, b.���� " & _
                  " From Ӧ����¼ a, ��Ӧ�� b " & _
                  " Where a.��λID = b.ID  And a.��¼����<>-1 and a.��¼״̬=[1] and a.ID=[2] Order by �ƻ����"
            Set rsInitCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint��¼״̬, mlngID)
            
            If rsInitCard.EOF Then
                If mEditType = gȡ�� Then
                    mErrBillStatusInfor = �Ѿ�����
                Else
                    mErrBillStatusInfor = 2
                End If
                Exit Sub
            End If
            
            txtNo.Caption = Nvl(rsInitCard!NO)
            mstrNo = txtNo
            
            If mEditType = gȡ�� Then
                Txt������ = gstrUserName
                Txt�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                Txt����� = gstrUserName
                Txt������� = Txt��������
            Else
                Txt������ = IIf(IsNull(rsInitCard!������), "", rsInitCard!������)
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End If
            If mEditType = g��� Then
                Txt����� = gstrUserName
                Txt������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            End If
            If mEditType = g�޸� Then
                Txt������ = gstrUserName
            End If
            txtEdit(13).Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mEditType = g�޸� Or mEditType = g���) And Nvl(rsInitCard!�����) <> "" Then
                mErrBillStatusInfor = 3
                Exit Sub
            End If
            txtEdit(0).Text = "[" & Nvl(rsInitCard!��Ӧ�̱���) & "]" & Nvl(rsInitCard!��Ӧ������)
            chkSelector.Tag = Nvl(rsInitCard!����)
            mlng��λID = Nvl(rsInitCard!��λID, 0)
            Dim intIndex As Integer
            Dim strTmp As String
            With rsInitCard
                For intIndex = 1 To 14
                    strTmp = txtEdit(intIndex).Tag
                    If InStr(1, strTmp, "���") <> 0 Then
                        txtEdit(intIndex).Text = Format(Nvl(.Fields(strTmp), 0), "####0.00;-####0.00; ;")
                    ElseIf strTmp = "�ɹ���" Or strTmp = "����" Then
                        txtEdit(intIndex).Text = Format(Nvl(.Fields(strTmp), 0), "####0.0000;-####0.0000; ;")
                    Else
                        txtEdit(intIndex).Text = Nvl(.Fields(strTmp))
                    End If
                Next
                If IsNull(!��Ʊ����) Then
                    Dtp��Ʊ����.Value = ""
                Else
                    Dtp��Ʊ����.Value = Format(!��Ʊ����, "yyyy-mm-dd")
                End If
            End With
            rsInitCard.Close
    End Select
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitControl()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����ؼ��е�����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    For intIndex = 1 To 14
         txtEdit(intIndex).Text = ""
    Next
    Dtp��Ʊ����.Value = ""
    txtNo = ""
End Sub

Private Sub chkSelector_Click()
    Call SetClass
End Sub

Private Sub chkSelector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    If mEditType = g�鿴 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mEditType = g��� Then        '���
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If ValidData = False Then Exit Sub
    
   If mEditType = gȡ�� Then
        If SaveStrike() = True Then
                Unload Me
        End If
        Exit Sub
    End If
    
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
            
        If IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mEditType = g�޸� Then    '�޸�
            Unload Me
            Exit Sub
        End If
        stbThis.Panels(2).Text = "��һ�ŵĵ��ݺţ�" & mstrNo
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    
    InitControl
    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
    mblnChange = False
End Sub

Private Sub cmdSelDept_Click()
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    strTemp = frm��Ӧ��ѡ��.SelDept(mstrPrivs)
    If strTemp = "" Then
        Unload frm��Ӧ��ѡ��
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Sub
    End If
    txtEdit(0).Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    mlng��λID = Val(Left(strTemp, InStr(strTemp, ",") - 1))
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord("select ���� from ��Ӧ�� where id=[1] ", Caption & "-��ȡ��Ӧ������", mlng��λID)
    If Not rsTemp.EOF Then
        chkSelector.Tag = Nvl(rsTemp!����)
    End If
    rsTemp.Close
    Call SetClass
    
    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdSelName_Click()
    Call GetItem("")
    txtEdit(2).SetFocus
End Sub

Private Sub Dtp��Ʊ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
     
    Call initCard
    Call SetEditPro
    Call chkSelector_Click
    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
    mblnChange = False
    setCtlEn
    Select Case mErrBillStatusInfor
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            ShowMsgbox "�õ����ѱ�ɾ�������飡"
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            ShowMsgbox "�õ����ѱ���������ˣ����飡"
            Unload Me
            Exit Sub
        Case �Ѿ�����
            '�޸ĵĵ����ѱ����
            ShowMsgbox "�õ���û�пɳ����ļ�¼�����飡"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mint��Ʊ��Len = Sys.FieldsLength("Ӧ����¼", "��Ʊ��")      'Get��Ʊ��Len
    txtEdit(8).MaxLength = mint��Ʊ��Len
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnYes As Boolean
    If mblnChange = False Then Exit Sub
    ShowMsgbox "���Ѿ������˵�����Ϣ,�������˳��Ļ�," & vbCrLf & "�����ĵ����ݽ����ܱ���,���Ҫ�˳���?", True, blnYes
    If blnYes = True Then Exit Sub
    SaveWinState Me, App.ProductName
    Cancel = 1
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub optClass_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 0 Then
        mlng��λID = 0
    End If
    setCtlEn
End Sub

Private Sub setCtlEn()
    Dim intIndex As Integer
    If mEditType = g��� Or mEditType = gȡ�� Or mEditType = g�鿴 Then
        Me.cmdOK.Enabled = True
    Else
        Me.cmdOK.Enabled = mblnChange
    End If
    chkSelector.Enabled = (mEditType = g���� Or mEditType = g�޸�)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Dim strTmp As String
    Dim blnOpen As Boolean
    
    strTmp = txtEdit(Index).Tag
    If InStr(1, strTmp, "���") <> 0 Or strTmp = "�ɹ���" Or strTmp = "����" Or InStr(1, strTmp, "��") <> 0 Then
            blnOpen = False
    Else
        blnOpen = True
    End If
    SetTxtGotFocus txtEdit(Index), blnOpen
End Sub

Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:��֤�Ϸ�,����True,����=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strTemp As String
    
   If mlng��λID = 0 Then
        ShowMsgbox "��Ӧ��ѡ������,������ѡ��!"
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Function
   End If
    
    For intIndex = 1 To 14
        strTemp = Trim(txtEdit(intIndex).Text)
        If intIndex = 1 Or txtEdit(intIndex).Tag = "��Ʊ���" Then
            If strTemp = "" Then
                ShowMsgbox txtEdit(intIndex).Tag & "��������!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        
        If strTemp <> "" Then
            If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(intIndex).MaxLength Then
                ShowMsgbox txtEdit(intIndex).Tag & "����,���������" & txtEdit(intIndex).MaxLength / 2 & "�����ֻ�" & txtEdit(intIndex).MaxLength & "���ַ�!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") <> 0 Then
                ShowMsgbox txtEdit(intIndex).Tag & "�������뵥����!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, txtEdit(intIndex).Tag, "���") <> 0 Or txtEdit(intIndex).Tag = "�ɹ���" Or txtEdit(intIndex).Tag = "����" Then
                If Not IsNumeric(strTemp) Then
                    ShowMsgbox txtEdit(intIndex).Tag & "����������,������!"
                    If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                    Exit Function
                End If
                If Val(strTemp) > 9999999999.99 Then
                    ShowMsgbox txtEdit(intIndex).Tag & "���ܴ���9999999999.99,������!"
                    If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                    Exit Function
                End If
                If Val(strTemp) < -999999999.99 Then
                    ShowMsgbox txtEdit(intIndex).Tag & "����С��-999999999.99,������!"
                    If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���濨Ƭ��Ϣ
    '--�����:
    '--������:
    '--��  ��:�ɹ�����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngID As Long
    Dim NO_IN As String
    
    On Error GoTo errHandle
    SaveCard = False
    
    If mEditType = g���� Then
        lngID = zlDatabase.GetNextId("Ӧ����¼")
        strSQL = "ZL_Ӧ����¼_INSERT("
        mstrNo = NextNo(67)
        NO_IN = mstrNo
    Else
        lngID = mlngID
        strSQL = "ZL_Ӧ����¼_UPDATE("
        NO_IN = Trim(txtNo)
    End If
    
    '���̲�������:
    '  Id_In         In Ӧ����¼.ID%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In         In Ӧ����¼.NO%Type,
    strSQL = strSQL & "'" & NO_IN & "',"
    '  �շ�id_In     In Ӧ����¼.�շ�id%Type,
    strSQL = strSQL & "" & "Null" & ","
    '  ��λid_In     In Ӧ����¼.��λid%Type,
    strSQL = strSQL & "" & mlng��λID & ","
    '  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(8).Text) = "", "NULL", "'" & Trim(txtEdit(8).Text) & "'") & ","
    '  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type,
    strSQL = strSQL & "" & IIf(Dtp��Ʊ����.Value = "" Or IsNull(Dtp��Ʊ����.Value), "NULL", "to_date('" & Format(Dtp��Ʊ����.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
    '  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type,
    strSQL = strSQL & "" & Val(txtEdit(9).Text) & ","
    '  �������_In   In Ӧ����¼.�������%Type,
    strSQL = strSQL & "" & "Null" & ","
    '  ��¼����_In   In Ӧ����¼.��¼����%Type,
    strSQL = strSQL & "" & "1" & ","
    '  ��ⵥ�ݺ�_In In Ӧ����¼.��ⵥ�ݺ�%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(5).Text) = "", "NULL", "'" & Trim(txtEdit(5).Text) & "'") & ","
    '  ���ݽ��_In   In Ӧ����¼.���ݽ��%Type,
    strSQL = strSQL & "" & Val(txtEdit(7).Text) & ","
    '  Ʒ��_In       In Ӧ����¼.Ʒ��%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(1).Text) = "", "NULL", "'" & Trim(txtEdit(1).Text) & "'") & ","
    '  ���_In       In Ӧ����¼.���%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(2).Text) = "", "NULL", "'" & Trim(txtEdit(2).Text) & "'") & ","
    '  ����_In       In Ӧ����¼.����%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(3).Text) = "", "NULL", "'" & Trim(txtEdit(3).Text) & "'") & ","
    '  ����_In       In Ӧ����¼.����%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(6).Text) = "", "NULL", "'" & Trim(txtEdit(6).Text) & "'") & ","
    '  ������λ_In   In Ӧ����¼.������λ%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(4).Text) = "", "NULL", "'" & Trim(txtEdit(4).Text) & "'") & ","
    '  ����_In       In Ӧ����¼.����%Type,
    strSQL = strSQL & "" & Val(txtEdit(10).Text) & ","
    '  �ɹ���_In     In Ӧ����¼.�ɹ���%Type,
    strSQL = strSQL & "" & Val(txtEdit(11).Text) & ","
    '  �ɹ����_In   In Ӧ����¼.�ɹ����%Type,
    strSQL = strSQL & "" & Val(txtEdit(12).Text) & ","
    '  ժҪ_In       In Ӧ����¼.ժҪ%Type,
    strSQL = strSQL & "" & IIf(Trim(txtEdit(13).Text) = "", "NULL", "'" & Trim(txtEdit(13).Text) & "'") & ","
    '  �������_In   In Ӧ����¼.�������%Type := Null
    strSQL = strSQL & "" & IIf(Trim(txtEdit(14).Text) = "", "NULL", "'" & Trim(txtEdit(14).Text) & "'") & ")"
 
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    mlngID = lngID
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SelMltProvide() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��Ӧ������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim strȨ�� As String
    
    If Trim(txtEdit(0).Text) = "" Then Exit Function
    
    strTmp = GetMatchingSting(UCase(txtEdit(0).Text), False)
    
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs)
    
    SelMltProvide = False
    
    strSQL = "" & _
        "  Select   ID,����,����,����,���֤��," & _
        "           to_char(���֤Ч��,'yyyy-mm-dd') as ���֤Ч��,ִ�պ�," & _
        "           to_char(ִ��Ч��,'yyyy-mm-dd') as ִ��Ч��,˰��ǼǺ�,��ϵ��,���� " & _
        "  From  ��Ӧ�� " & _
        "  Where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
        "       " & zl_��ȡվ������ & "  and ĩ��=1  " & _
        "       And ( ���� Like [1] or ���� like [1] or ����  like upper([1])) " & strȨ��
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp)
    
    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ�ָ���Ĺ�Ӧ��!"
        Exit Function
    End If
    With rsTemp
        If .RecordCount > 1 Then
            mstrSelectTag = "Provide"
            Set mshSelect.Recordset = rsTemp
            With mshSelect
                .Top = txtEdit(0).Top + txtEdit(0).Height + 10
                .Left = txtEdit(0).Left
                .Visible = True
                .ColWidth(0) = 0
                .ColWidth(1) = 1400
                .ColWidth(2) = 2000
                .ColWidth(3) = 800
                .ColWidth(5) = 1000
                .ColWidth(6) = 1400
                .ColWidth(7) = 1000
                .ColWidth(8) = 1400
                .ColWidth(9) = 1000
                .ColWidth(10) = 0
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder
                .SetFocus
                Exit Function
            End With
        Else
            txtEdit(0).Text = "[" & Nvl(rsTemp!����) & "]" & rsTemp!����
            mlng��λID = Nvl(rsTemp!ID, 0)
            SelMltProvide = True
        End If
    End With
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            If SelMltProvide = False And mshSelect.Visible = False Then
                If txtEdit(0).Enabled Then txtEdit(0).SetFocus
            Else
                If mshSelect.Visible = False Then
                    zlCommFun.PressKey vbKeyTab
                End If
            End If
        ElseIf Index = 1 And chkSelector.Value = 1 Then
            If Trim(txtEdit(Index)) <> "" And txtEdit(Index).Tag <> txtEdit(Index).Text Then
                GetItem UCase(Trim(txtEdit(Index)))
            End If
            zlCommFun.PressKey vbKeyTab
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strTmp As String
    
    strTmp = txtEdit(Index).Tag
    
    If InStr(1, strTmp, "���") <> 0 Or strTmp = "�ɹ���" Or strTmp = "����" Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�����ʽ
    Else
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Dim strTmp As String
    
    strTmp = txtEdit(Index).Tag
    
    If InStr(1, strTmp, "���") <> 0 Then
        txtEdit(Index) = Format(Val(txtEdit(Index).Text), "####0.00;-####0.00; ;")
        If strTmp = "�ɹ����" Then
            txtEdit(11) = Format(Val(txtEdit(Index).Text) / IIf(Val(txtEdit(10)) = 0, 1, Val(txtEdit(10))), "####0.0000;-####0.0000; ;")
        End If
    ElseIf strTmp = "�ɹ���" Then
        txtEdit(Index) = Format(Val(txtEdit(Index).Text), "####0.0000;-####0.0000; ;")
        txtEdit(12) = Format(Val(txtEdit(Index).Text) * Val(txtEdit(10).Text), "####0.00;-####0.00; ;")
    ElseIf strTmp = "����" Then
        txtEdit(Index) = Format(Val(txtEdit(Index).Text), "####0.0000;-####0.0000; ;")
        txtEdit(12) = Format(Val(txtEdit(Index).Text) * Val(txtEdit(11).Text), "####0.00;-####0.00; ;")
    ElseIf strTmp = "��ⵥ�ݺ�" Then
        Dim intYear  As Integer, strYear As String
        If IsNumeric(txtEdit(Index).Text) And txtEdit(Index).Text <> "" Then
            If Len(txtEdit(Index).Text) < 8 And Len(txtEdit(Index).Text) > 0 Then
                txtEdit(Index).Text = UCase(LTrim(txtEdit(Index).Text))
                intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
                strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
                txtEdit(Index).Text = strYear & String(7 - Len(txtEdit(Index).Text), "0") & txtEdit(Index).Text
            End If
        End If
    End If
    
    ImeLanguage False
End Sub

Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��˵���
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    '   ZL_Ӧ����¼_Verify���̲���:
    '    ID_IN
    
    On Error GoTo errHandle:
    
    gstrSQL = "ZL_Ӧ����¼_Verify(" & _
        mlngID & ")"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
 '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    
    
    '   ZL_Ӧ����¼_STRIKE���̲���:
    '    ID_IN
    
    On Error GoTo errHandle:
    
    gstrSQL = "ZL_Ӧ����¼_STRIKE(" & _
        mlngID & ")"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveStrike = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub mshselect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sinWidth As Single
    
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sinWidth = sinWidth + .ColWidth(i)
                    If sinWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    With mshSelect
        Select Case mstrSelectTag
            Case "Provide"
                If KeyAscii = vbKeyReturn Then
                    If .Row = 0 Then Exit Sub
                    txtEdit(0).Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
                    
                    chkSelector.Tag = .TextMatrix(.Row, 10)
                    Call SetClass
                    
                    mlng��λID = Val(.TextMatrix(.Row, 0))
                    If txtEdit(1).Enabled Then txtEdit(1).SetFocus
                ElseIf KeyAscii = 27 Then
                    If txtEdit(0).Enabled Then txtEdit(0).SetFocus
                End If
            Case Else
        End Select
        .Visible = False
    End With
End Sub

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ñ༭����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    For intIndex = 0 To 14
        txtEdit(intIndex).Enabled = mblnEdit
    Next
    cmdSelDept.Enabled = mblnEdit
    Dtp��Ʊ����.Enabled = mblnEdit
End Sub
'��ӡ����
Private Sub printbill()
    ReportOpen gcnOracle, glngSys, "ZL1_bill_1322", Me, "ID=" & mlngID
End Sub

Private Function GetClassValue() As Integer
    Dim i As Integer
    For i = 0 To optClass.Count - 1
        If optClass(i).Value And optClass(i).Enabled Then
            GetClassValue = i
            Exit Function
        End If
    Next
    GetClassValue = -1
End Function

Private Sub GetItem(ByVal strKey As String)
    Dim intClass As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim sngX As Single, sngY As Single, sngH As Single
    Dim intSysParam As Integer
    Dim strMatch As String
    
    intClass = GetClassValue()
    vRect = zlControl.GetControlRect(txtEdit(1).hwnd)
    sngX = vRect.Left
    sngY = vRect.Bottom
    
    On Error GoTo errHandle
    Select Case intClass
    Case 0
        'ҩƷ
        If strKey = "" Then
            strSQL = "Select ID, �ϼ�id, ����, ����, '' ���, '' ����, '' ҩ�ⵥλ, '' סԺ��λ, '' ���ﵥλ, 0 As ĩ�� " & _
                     "From ���Ʒ���Ŀ¼ " & vbLf & _
                     "Where ���� in ('1','2','3') " & vbLf & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select a.Id, c.����id As �ϼ�id, a.����, a.����, a.���, a.����, b.ҩ�ⵥλ, b.סԺ��λ, b.���ﵥλ, 1 As ĩ�� " & vbLf & _
                     "From �շ���ĿĿ¼ A, ҩƷ��� B, ������ĿĿ¼ C " & vbLf & _
                     "Where a.Id = b.ҩƷid And b.ҩ��id = c.Id And a.��� in ('5','6','7') " & vbLf & _
                     "  And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-ҩƷ" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct a.ID, null �ϼ�ID, a.����, a.����, a.���, a.����, b.ҩ�ⵥλ, b.סԺ��λ, b.���ﵥλ " & vbLf & _
                     "From �շ���ĿĿ¼ A, ҩƷ��� B, �շ���Ŀ���� C " & vbLf & _
                     "Where a.Id = b.ҩƷid And a.id = c.�շ�ϸĿid And A.��� in ('5','6','7') " & vbLf & _
                     "  And (to_char(A.����ʱ��, 'yyyy-mm-dd') = '3000-01-01' or A.����ʱ�� is null) " & _
                     "  And C.���� = 1 "
            intSysParam = Val(zlDatabase.GetPara("���뷽ʽ"))
            strMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (a.���� Like [1] Or C.���� Like [2] And C.����=3) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And C.���� Like [2] and c.����=" & IIf(intSysParam = 0, 1, 2) & " "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And C.���� Like [2] "
            Else
                strSQL = strSQL & " And (a.���� = [1] And C.���� Like [2] Or C.���� LIKE [2]) and c.����=" & IIf(intSysParam = 0, 1, 2) & " "
            End If
            strSQL = strSQL & vbNewLine & "Order by a.���� "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-ҩƷ" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , strMatch & strKey & "%")
        End If
        
    Case 1
        '����
        If strKey = "" Then
            strSQL = "Select ID, �ϼ�id, ����, ����, '' ���, '' ����, '' As ���㵥λ, 0 As ĩ�� " & _
                     "From ���Ʒ���Ŀ¼ " & vbLf & _
                     "Where ���� = '7' " & vbLf & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & vbLf & _
                     "Union all " & vbLf & _
                     "Select i.Id, b.����id As �ϼ�id, i.����, i.����, i.���, i.����, i.���㵥λ, 1 As ĩ�� " & vbLf & _
                     "From �շ���ĿĿ¼ I, �������� T, ������ĿĿ¼ B " & vbLf & _
                     "Where i.Id = t.����id And t.����id = b.Id And i.��� = '4' " & vbLf & _
                     "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-����" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select Distinct i.Id, i.����, i.����, i.���, i.����, i.���㵥λ, 1 As ĩ�� " & vbLf & _
                     "From �շ���ĿĿ¼ I, �������� T, �շ���Ŀ���� B " & vbLf & _
                     "Where i.Id = t.����id And i.Id = b.�շ�ϸĿid And i.��� = '4' " & vbLf & _
                     "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
            intSysParam = Val(zlDatabase.GetPara("���뷽ʽ"))
            strMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (i.���� Like [1] Or b.���� Like [2] And b.����=3) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And b.���� Like [2] And b.���� = [3] "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And b.���� Like [2] "
            Else
                strSQL = strSQL & " And (i.���� = [1] And b.���� Like [2] Or b.���� LIKE [2]) And b.���� = [3] "
            End If
            strSQL = strSQL & vbLf & "Order by i.���� "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-����" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , strMatch & strKey & "%" _
                    , IIf(intSysParam = 0, 1, 2))
        End If
    Case 2
        '����
        If strKey = "" Then
            strSQL = "Select ID, 0 ĩ��, �ϼ�id, ����, ����, '' ���, '' ����, '' ɢװ��λ, '' ��װ��λ " & _
                     "From ���ʷ��� " & _
                     "Where ������� in ('��ͨ����', 'ҽ������') " & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & _
                     "Union All " & _
                     "Select ID, 1 ĩ��, ����id �ϼ�id, ����, ����, ���, ����, ɢװ��λ, ��װ��λ " & _
                     "From ����Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) And ������� in ('��ͨ����', 'ҽ������') "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-����" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, ����, ����, ���, ����, ɢװ��λ, ��װ��λ " & _
                     "From ����Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) And ������� in ('��ͨ����', 'ҽ������') "
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (���� Like [1] Or ���� Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And ���� Like [2] "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And ���� Like [2] "
            Else
                strSQL = strSQL & " And (���� = [1] And ���� Like [2] Or ���� LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by ���� "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-����" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , "%" & strKey & "%")
        End If
    Case 3
        '�豸
        If strKey = "" Then
            strSQL = "Select ID, 0 ĩ��, �ϼ�id, ����, ����, '' ���, '' ����, '' ��λ " & _
                     "From �豸���� " & _
                     "Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & _
                     "Union All " & _
                     "Select ID, 1 ĩ��, ����id �ϼ�id, ����, ����, ���, ����, ��λ " & _
                     "From �豸Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) "
            
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, Caption & "-�豸" _
                    , False, "", "ѡ��", False, False, False, sngX, sngY, sngH, blnCancel, False, False)
        Else
            strSQL = "Select ID, ����, ����, ���, ����, ��λ " & _
                     "From �豸Ŀ¼ " & _
                     "Where (to_char(����ʱ��,'yyyy-MM-DD') = '3000-01-01' or ����ʱ�� is null) "
            
            If IsNumeric(strKey) Then
                strSQL = strSQL & " And (���� Like [1] Or ���� Like [2]) "
            ElseIf zlCommFun.IsCharAlpha(strKey) Then
                strSQL = strSQL & " And ���� Like [2] "
            ElseIf zlCommFun.IsCharChinese(strKey) Then
                strSQL = strSQL & " And ���� Like [2] "
            Else
                strSQL = strSQL & " And (���� = [1] And ���� Like [2] Or ���� LIKE [2]) "
            End If
            strSQL = strSQL & vbLf & "Order by ���� "
        
            Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Caption & "-�豸" _
                    , False, "", "ѡ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False _
                    , strKey & "%" _
                    , "%" & strKey & "%")
        End If
    End Select
    
    If blnCancel = False And Not rsTemp Is Nothing Then
        txtEdit(1).Text = Nvl(rsTemp!����)
        txtEdit(1).Tag = Nvl(rsTemp!����)
        txtEdit(2).Text = Nvl(rsTemp!���)
        txtEdit(3).Text = Nvl(rsTemp!����)
        If intClass = 1 Then
            txtEdit(4).Text = Nvl(rsTemp!���㵥λ)
        ElseIf intClass = 3 Then
            txtEdit(4).Text = Nvl(rsTemp!��λ)
        End If
    End If
    If Not rsTemp Is Nothing Then rsTemp.Close
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Sub SetClass()
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    For i = 0 To optClass.Count - 1
        'ϵͳ
        If i >= 2 Then
            Set rsTemp = zlDatabase.OpenSQLRecord("Select Count(1) Rec From zlSystems Where ��� = [1]", Caption, IIf(i = 2, 400, 600))
            optClass(i).Enabled = rsTemp!rec > 0
            rsTemp.Close
        Else
            optClass(i).Enabled = True
        End If
        'Ȩ��
        Select Case i
            Case 0
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";ҩƷ;") > 0
            Case 1
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";����;") > 0
            Case 2
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";����;") > 0
            Case 3
                optClass(i).Enabled = optClass(i).Enabled And InStr(mstrPrivs, ";�豸;") > 0
        End Select
        optClass(i).Visible = chkSelector.Value = 1
    Next
    '��Ӧ��
    If Len(chkSelector.Tag) >= 1 Then  'ҩƷ
        optClass(0).Enabled = optClass(0).Enabled And Mid(chkSelector.Tag, 1, 1) = "1"
    Else
        optClass(0).Enabled = False
    End If
    If Len(chkSelector.Tag) >= 5 Then  '����
        optClass(1).Enabled = optClass(1).Enabled And Mid(chkSelector.Tag, 5, 1) = "1"
    Else
        optClass(1).Enabled = False
    End If
    If Len(chkSelector.Tag) >= 2 Then  '����
        optClass(2).Enabled = optClass(2).Enabled And Mid(chkSelector.Tag, 2, 1) = "1"
    Else
        optClass(2).Enabled = False
    End If
    If Len(chkSelector.Tag) >= 3 Then  '�豸
        optClass(3).Enabled = optClass(3).Enabled And Mid(chkSelector.Tag, 3, 1) = "1"
    Else
        optClass(3).Enabled = False
    End If
    
    cmdSelName.Visible = chkSelector.Value = 1
    If chkSelector.Value = 1 Then
        txtEdit(1).Width = txtEdit(2).Width - cmdSelName.Width
    Else
        txtEdit(1).Width = txtEdit(2).Width
    End If
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub
