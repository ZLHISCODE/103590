VERSION 5.00
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#2.0#0"; "zlIDKind.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSurety 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "סԺ������Ϣ����"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmSurety.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7575
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicDeposit 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   720
      ScaleHeight     =   3090
      ScaleWidth      =   5790
      TabIndex        =   18
      Top             =   3360
      Width           =   5790
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   2265
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���ܶ"
         Height          =   180
         Left            =   45
         TabIndex        =   19
         Top             =   75
         Width           =   900
      End
   End
   Begin VB.Frame fraPati 
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   7350
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1350
         TabIndex        =   3
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2640
         Picture         =   "frmSurety.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����(F2)"
         Top             =   225
         Width           =   300
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         ToolTipText     =   "��ݼ�F4"
         Top             =   225
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   $"frmSurety.frx":0914
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lblCur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ��ʽ��"
         Height          =   180
         Left            =   5085
         TabIndex        =   22
         Top             =   285
         Width           =   1260
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�ȼ���"
         Height          =   180
         Left            =   5085
         TabIndex        =   21
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2985
         TabIndex        =   5
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Left            =   3960
         TabIndex        =   6
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�ţ�"
         Height          =   180
         Left            =   330
         TabIndex        =   7
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ң�"
         Height          =   180
         Left            =   2325
         TabIndex        =   8
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lblBed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ţ�"
         Height          =   180
         Left            =   3960
         TabIndex        =   9
         Top             =   630
         Width           =   540
      End
   End
   Begin VB.PictureBox picSurety 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   120
      ScaleHeight     =   3900
      ScaleWidth      =   7425
      TabIndex        =   16
      Top             =   1080
      Width           =   7425
      Begin VB.Frame fraEdit 
         Caption         =   "��Ϣ����"
         Height          =   1095
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   7335
         Begin VB.TextBox txt������ 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   9
            TabIndex        =   28
            Top             =   360
            Width           =   1005
         End
         Begin VB.TextBox txt������ 
            Height          =   300
            Left            =   840
            MaxLength       =   100
            TabIndex        =   27
            Top             =   360
            Width           =   1005
         End
         Begin VB.CheckBox chkUnlimit 
            Caption         =   "���޶��"
            Height          =   255
            Left            =   2760
            TabIndex        =   26
            ToolTipText     =   "���޵�����ʱ�������õ���ʱ��"
            Top             =   720
            Width           =   1050
         End
         Begin VB.CheckBox chk��ʱ���� 
            Caption         =   "��ʱ����"
            Height          =   255
            Left            =   840
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtReason 
            Height          =   300
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   24
            Top             =   720
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   300
            Left            =   5040
            TabIndex        =   29
            Top             =   360
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   190906371
            CurrentDate     =   38915.6041666667
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   2160
            TabIndex        =   33
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl������ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   240
            TabIndex        =   32
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl����ʱ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   4140
            TabIndex        =   31
            ToolTipText     =   "��Ժ���˲���ʹ��ʱ�޵���"
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��"
            Height          =   180
            Left            =   4140
            TabIndex        =   30
            Top             =   780
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdAdd 
         Cancel          =   -1  'True
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "�������һ��������¼���ڻ�û����������ʱ����������"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&M)"
         Height          =   350
         Left            =   1350
         TabIndex        =   11
         ToolTipText     =   "ֻ�����޸����һ��������¼"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   2450
         TabIndex        =   12
         ToolTipText     =   "ֻ����ɾ�����һ��������¼"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   6000
         TabIndex        =   13
         ToolTipText     =   "(F9)�˳�"
         Top             =   1200
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
         Height          =   2265
         Left            =   0
         TabIndex        =   14
         Top             =   1680
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   5145
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
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
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   3945
      Left            =   555
      TabIndex        =   17
      Top             =   1035
      Width           =   3795
      _Version        =   589884
      _ExtentX        =   6694
      _ExtentY        =   6959
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmSurety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlng����ID As Long
Private mbln��Ժ���� As Boolean
Private mstrPrivs As String
Private mlng��ҳID As Long      '��Ժ����Ϊ��ǰסԺ�Ǽǵ���ҳID
Private mlngModul As Long       '0-������Ϣ����;1-������Ժ����

'����
Private mblnSeekName As Boolean         '����ģ������
Private mblnCheckPass As Boolean        '��Ժ�Ǽ�ʱˢ����������


Private mrsInfo As New ADODB.Recordset
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjOneCardCom As Object
Private mblnDefaultPassInputCardNo As Boolean 'ȱʡˢ���Ƿ��������뿨��
Private mblnNotClick As Boolean
Private mblnFirst As Boolean
Private mstr�������� As String

Private Enum E_COL_Deposit
    COL_���� = 1
    COL_���ݺ� = 2
    COL_����
    COL_���
    COL_�ɿ���
    COL_����
    COL_�տ���
End Enum

Private Enum E_COL_Surety
    COL_��� = 1
    COL_������
    COL_������
    COL_��ʱ����
    COL_����ԭ��
    COL_�Ǽ�ʱ��
    COL_����ʱ��
    COL_ɾ����־
    COL_����Ա����
    COL_����Ա���
    COL_ɾ������Ա����
    COL_ɾ������Ա���
    COL_ɾ��ʱ��
End Enum
 
Public Function ShowMe(ByRef frmMain As Form, ByVal lngPatiid As Long, ByVal bln��Ժ���� As Boolean, ByVal strPrivs As String, _
    ByVal lngModul As Long, Optional blnSeekName As Boolean, Optional blnCheckPass As Boolean) As Boolean
'����:������Ϣ����
'   frmMain-���������
'   lngPatiId-����ID
'   bln��Ժ����-T-��Ժ����
'   strPrivs-Ȩ��
'   lngModul= 1101-������Ϣ����,1131-P������Ժ����
'   blnSeekName-������ģ������(ģ��=������Ժ����ʱ�����ֵ)
    mlng����ID = lngPatiid
    mbln��Ժ���� = bln��Ժ����
    mstrPrivs = strPrivs
    mlngModul = lngModul
    
    If mlngModul = P������Ժ���� Then
        mblnSeekName = blnSeekName  '
        mblnCheckPass = blnCheckPass
    End If
    
    Me.Show 1, frmMain
    ShowMe = True
End Function
  
Private Sub chkUnlimit_Click()
     '���޵����������ʱ����
    If chkUnlimit.Value = 1 And IsNull(dtp����ʱ��.Value) Then
        dtp����ʱ��.Value = DateAdd("d", 3, dtp����ʱ��.MinDate)
    End If
    chk��ʱ����.Enabled = Not (chkUnlimit.Value = 1)
    txt������.Enabled = Not (chkUnlimit.Value = 1)
    
    If chkUnlimit.Value = 1 Then
        txt������.Text = "999999999":  txt������.BackColor = vbInactiveCaptionText
    Else
        txt������.Text = "": txt������.BackColor = vbWhite
    End If
End Sub

Private Sub chk��ʱ����_Click()
    If chk��ʱ����.Value = 1 Then
        dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
        dtp����ʱ��.Value = Null
        chkUnlimit.Value = 0        'ֵ�ı�ʱ����ʽ����click�¼�
    End If
    chkUnlimit.Enabled = Not (chk��ʱ����.Value = 1) And mbln��Ժ����
    dtp����ʱ��.Enabled = Not (chk��ʱ����.Value = 1) And mbln��Ժ����
End Sub

Private Sub cmdDel_Click()
    Dim strJsonIn As String, strJsonOut As String
    Dim str�Ǽ�ʱ�� As String
    Dim strɾ����־ As String
    Dim blnOk As Boolean
    
    If mlngModul = P������Ժ���� Then
        blnOk = True
        If mrsInfo Is Nothing Then
            blnOk = False
        ElseIf mrsInfo.State = adStateClosed Then
            blnOk = False
        End If
        
        If blnOk = False Then
            stbThis.Panels(1).Text = "û��ȷ��Ҫ���е����Ĳ���!"
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
        End If
    End If
    '����21368 by lesfeng 2010-08-02
    strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
    If strɾ����־ = "ɾ��" Then
        MsgBox "����������¼�Ѿ�Ϊɾ����ǣ����ܽ���ɾ����ǲ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪ���б�Ǵ���������¼Ϊɾ�������?" & vbCrLf & vbCrLf & "ע��,ɾ����Ǻ󣬵�ǰ�������᲻�ָܻ�!" _
        , vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    
    If Trim(msh.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��"))) = "" Then
        str�Ǽ�ʱ�� = ""
    Else
        str�Ǽ�ʱ�� = Format(Trim(msh.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��"))), "YYYY-MM-DD HH:MM:SS")
    End If
    '����21368 by lesfeng 2010-08-02
    strJsonIn = GetJsonSurety(3, mlng����ID, mlng��ҳID, , , , , , str�Ǽ�ʱ��)
    If strJsonIn <> "" Then
        Call CallService("Zl_Exsesvr_Updatepatisurety", strJsonIn, strJsonOut, "ɾ��������¼", P������Ժ����, False, , , , True)
    End If
    
    stbThis.Panels(1).Text = "ɾ�������ɹ�!"
    Call LoadSurety
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdModify_Click()
    Dim strSQL As String, str������ As String, str����ʱ�� As String
    Dim str�Ǽ�ʱ�� As String
    Dim strɾ����־ As String
    Dim blnOk As Boolean
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    'ֻ���޸ĵ�ǰѡ�в�����Ч�ĵ�����¼
    
    If cmdModify.Caption = "�޸�(&M)" Then
        If mlngModul = P������Ժ���� Then
            blnOk = True
            If mrsInfo Is Nothing Then
                blnOk = False
            ElseIf mrsInfo.State = adStateClosed Then
                blnOk = False
            End If
            
            If blnOk = False Then
                stbThis.Panels(1).Text = "û��ȷ��Ҫ���е����Ĳ���!"
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                Exit Sub
            End If
        End If
    
        '��ȡ�޸���Ϣ
        If msh.TextMatrix(msh.Row, GetColNum("������")) = "" Then
            stbThis.Panels(1).Text = "û�п����޸ĵĵ�����Ϣ!"
            Exit Sub
        End If
        '����21368 by lesfeng 2010-08-02
        strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
        If strɾ����־ = "ɾ��" Then
            MsgBox "����������¼�Ѿ�Ϊɾ����ǣ����ܽ����޸Ĳ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        cmdModify.Caption = "����(&S)"
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        cmdExit.Caption = "ȡ��(&C)"
        fraEdit.Enabled = True
        
        With msh
            txt������.Text = Trim(.TextMatrix(.Row, GetColNum("������")))
            If .TextMatrix(.Row, GetColNum("������")) = "����" Then
                chkUnlimit.Value = 1    'ֵ��ͬʱ��ʽ����click�¼�
                txt������.Text = "999999999"
            Else
                chkUnlimit.Value = 0
                txt������.Text = Val(.TextMatrix(.Row, GetColNum("������")))
            End If
            
            If IsDate(.TextMatrix(.Row, GetColNum("����ʱ��"))) Then
                dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm"
                dtp����ʱ��.Value = CDate(.TextMatrix(.Row, GetColNum("����ʱ��")))
            Else
                dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm" '������ɼ��������ִ�л����
                dtp����ʱ��.Value = Null
            End If
            
            chk��ʱ����.Value = IIf(.TextMatrix(.Row, GetColNum("��ʱ����")) = "��", 1, 0)
            If txt������.Enabled Then txt������.SetFocus
            txt������.Tag = Trim(.TextMatrix(msh.Row, GetColNum("�Ǽ�ʱ��")))
        End With
    Else
    '�����޸Ľ��
        '1.���ݼ��
        If Not Check������Ϣ Then Exit Sub
        '�Ȼָ����水ť״̬
        cmdModify.Caption = "�޸�(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "�˳�(&X)"
        fraEdit.Enabled = True      'SetCanEdit���ٴ�����
        
        str������ = Replace(Trim(txt������.Text), "'", "''")
        str����ʱ�� = ""
        If Not IsNull(dtp����ʱ��.Value) Then str����ʱ�� = Format(dtp����ʱ��.Value, "YYYY-MM-DD hh:mm:ss")
        str�Ǽ�ʱ�� = Format(txt������.Tag, "YYYY-MM-DD hh:mm:ss")
        
        '���ȼ��
        If Not CheckLen(txt������, 64) Then Exit Sub
        
        '2.���ݱ���
        On Error GoTo errH
        strJsonIn = GetJsonSurety(2, mlng����ID, mlng��ҳID, str������, Val(txt������.Text), chk��ʱ����.Value, Trim(txtReason.Text), str����ʱ��, str�Ǽ�ʱ��)
        If strJsonIn <> "" Then
            Call CallService("Zl_Exsesvr_Updatepatisurety", strJsonIn, strJsonOut, "���µ�����¼", P������Ժ����, False, , , , True)
        End If
    
        '3.����ˢ��
        stbThis.Panels(1).Text = "�޸Ľ���ѱ���!"
        Call LoadSurety
        Call Init������Ϣ
        If cmdExit.Enabled Then cmdExit.SetFocus
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init������Ϣ()
    Dim Datsys As Date

    txt������.Text = ""
    chkUnlimit.Enabled = mbln��Ժ����
    chkUnlimit.Value = 0            '���ֵ�б仯,����ʽ����click�¼�
    txt������.Text = ""
    txtReason.Text = ""
    
    dtp����ʱ��.Enabled = mbln��Ժ����
    dtp����ʱ��.CheckBox = True: dtp����ʱ��.CustomFormat = "yyyy-MM-dd HH:mm" '����checkbox�ɼ���
    If dtp����ʱ��.Enabled Then
        Datsys = zldatabase.Currentdate
        dtp����ʱ��.MinDate = Datsys
        dtp����ʱ��.Value = DateAdd("d", 3, Datsys)
    End If
    dtp����ʱ��.Value = Null
    
    chk��ʱ����.Enabled = True
    chk��ʱ����.Value = 0
    chkUnlimit.TabStop = True
End Sub

Public Sub InitFace()
    lblSex.Caption = "�Ա�": lblNO.Caption = "סԺ�ţ�": lblBed.Caption = "���ţ�"
    lblAge.Caption = "���䣺": lblDept.Caption = "���ң�": lblDeposit.Caption = "Ԥ���ܶ"
    lblType.Caption = "�ѱ�ȼ���": lblCur.Caption = "ҽ�Ƹ��ʽ��"
End Sub

Private Sub cmdPati_Click()
    If frmPatiSelect.ShowMe(Me) = True Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub dtp����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = vbKeySpace Then
        If dtp����ʱ��.CheckBox Then
            KeyAscii = 0
            If IsNull(dtp����ʱ��.Value) Then
                dtp����ʱ��.Value = DateAdd("d", 3, zldatabase.Currentdate)
            Else
                dtp����ʱ��.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mlngModul = P������Ժ���� Then
        If mblnFirst = True Then Exit Sub
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        mblnFirst = True
    End If
End Sub

Private Sub Form_Load()
        
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    Me.Height = 5925
    If mlngModul = P������Ϣ���� Then
        Call LoadSurety
        Call Init������Ϣ
        
        If InStr(mstrPrivs, "������Ϣ����") <= 0 Then
            cmdAdd.Visible = False
        End If
        
        If InStr(mstrPrivs, "������Ϣ����") <= 0 Then
            cmdModify.Visible = False
        End If
        
        If InStr(mstrPrivs, "������Ϣɾ��") <= 0 Then
            cmdDel.Visible = False
        End If
        
        If InStr(mstrPrivs, "������Ϣ����") <= 0 And InStr(mstrPrivs, "������Ϣ����") And InStr(mstrPrivs, "������Ϣɾ��") <= 0 Then
            fraEdit.Enabled = False
            Me.Caption = "������Ϣ�鿴(��ǰ�û���" & UserInfo.���� & ")"
        End If
        
        Me.Caption = "������Ϣ����"
        Me.fraPati.Visible = False
        Me.PicDeposit.Visible = False
        Me.Height = Me.Height - fraPati.Height
        Me.picSurety.Top = fraPati.Top
        Me.tbcPage.Visible = False
    Else
        mblnFirst = False
        Call InitTabPage
        Call zlCardSquareObject
        Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, mobjOneCardCom, "", txtPatient)
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.hWnd)
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
        IDKind.Enabled = True
        
        If Not mobjOneCardCom Is Nothing Then
            IDKind.IDKindStr = mobjOneCardCom.zlGetIDKindStr(IDKind.IDKindStr)
        End If
        
        Call ClearWinInfor(True)
        
        fraEdit.Enabled = True
        If InStr(mstrPrivs, "����Ǽ�") <= 0 And InStr(mstrPrivs, "����ԤԼ") = 0 And InStr(mstrPrivs, "���ղ��˵Ǽ�") <= 0 Then
            fraEdit.Enabled = False
            cmdAdd.Visible = False
            cmdModify.Visible = False
            cmdDel.Visible = False
            Me.Caption = "סԺ������Ϣ�鿴(��ǰ�û���" & UserInfo.���� & ")"
        End If
        txt������.Enabled = fraEdit.Enabled
        txt������.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
        txt������.Enabled = fraEdit.Enabled
        txt������.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
        chk��ʱ����.Enabled = fraEdit.Enabled
        chkUnlimit.Enabled = fraEdit.Enabled
        txtReason.Enabled = fraEdit.Enabled
        txtReason.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
        If mlng����ID > 0 Then
            txtPatient.Text = "-" & mlng����ID
            Call txtPatient_KeyPress(vbKeyReturn)
        Else
            cmdAdd.Enabled = False
        End If
        Me.Caption = "סԺ������Ϣ����"
    End If
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
End Sub

Private Sub ClearWinInfor(Optional ByVal blnClear As Boolean = False)
    Call InitFace
    Call LoadSurety(blnClear)
    Call LoadPrepay(blnClear)
    Call Init������Ϣ
End Sub

Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
        
    Set objItem = tbcPage.InsertItem(1, "������Ϣ", picSurety.hWnd, 0)
    objItem.Tag = 1
    
    Set objItem = tbcPage.InsertItem(2, "Ԥ����Ϣ", PicDeposit.hWnd, 0)
    objItem.Tag = 2
    
    With tbcPage
        .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Function GetColNumList(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNumList = i: Exit Function
    Next
    GetColNumList = -1
End Function

Private Sub SetSuretyHeader()
    Dim strHead As String, i As Long
    strHead = ",4,300|���,4,1000|������,4,800|������,7,1250|��ʱ����,4,850|����ԭ��,4,1800|�Ǽ�ʱ��,1,1800|����ʱ��,1,1800|ɾ����־,4,850|����Ա����,4,1050|����Ա���,4,1050|ɾ������Ա����,4,1050|ɾ������Ա���,4,1050|ɾ��ʱ��,1,1800"
    With msh
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        .FixedRows = 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub SetDepositHeader()
    Dim strHead As String, i As Long
    strHead = ",4,300|����,4,1350|���ݺ�,4,1110|����,1,1200|���,1,0|�ɿ���,7,1600|����,4,1000|�տ���,1,1000"
    With mshList
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub GetSuretyBalance()
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim dblMoney As Double
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim colPati As Collection
    
    On Error GoTo errH
    If mlngModul = P������Ϣ���� Then
        If zl_PatiSvr_GetPatiInfo(mlng����ID, Nothing, colPati, 2) Then
            If colPati(1)("_pati_dept_id") = 0 Then
                mlng��ҳID = 0
            Else
                mlng��ҳID = colPati(1)("_pati_pageid")
            End If
        End If
    End If
    strJsonIn = GetNode("pati_id", mlng����ID, True)
    strJsonIn = strJsonIn & GetNode("pati_pageid", mlng��ҳID)
    strJsonIn = GetNode("input", "{" & strJsonIn & "}", True, 1)
    If CallService("Zl_Exsesvr_Getpatisurety", strJsonIn, strJsonOut, "��ȡ��Ч������", mlngModul) Then
        dblMoney = Val(GetJsonNodeValue("output.guarantee_money") & "")
        If dblMoney > 0 Then
            stbThis.Panels(2).Text = "��Ч������:" & dblMoney
        Else
            stbThis.Panels(2).Text = "��Ч������:��"
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSurety(Optional ByVal blnClear As Boolean = False)
    Dim Curdate As Date
    Dim i As Integer, lngRow As Integer, RowPageid As Integer
    Dim strɾ����־ As String
    Dim lng����ID As Long, lng��ҳid As Long
    Dim collist As Collection
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    On Error GoTo errH
    If mrsInfo Is Nothing Then
        lng����ID = mlng����ID
        lng��ҳid = mlng��ҳID
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = mlng����ID
        lng��ҳid = mlng��ҳID
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        lng��ҳid = Val(Nvl(mrsInfo!��ҳID))
    End If
    stbThis.Panels(2).Text = ""
    If blnClear = True Then
        msh.Rows = 1
        msh.Rows = 2
        msh.RowData(1) = 0
        Call SetSuretyHeader
    Else
        Call SetSuretyHeader
        strJsonIn = GetNode("pati_id", lng����ID, True)
        If mlngModul = P������Ժ���� Then strJsonIn = strJsonIn & GetNode("pati_pageid", lng��ҳid)
        strJsonIn = GetNode("input", "{" & strJsonIn & "}", True, 1)
        
        If CallService("Zl_Exsesvr_Getpatisuretylist", strJsonIn, strJsonOut, "��ȡ������¼", mlngModul) Then
            Set collist = GetJsonListValue("output.item_list")
        Else
            Set collist = New Collection
        End If
        If collist.Count > 0 Then
            With msh
                .Redraw = False
                .Rows = collist.Count + 1
                .FixedRows = 1
                For i = 1 To collist.Count
                    .RowData(i) = lng����ID
                    .TextMatrix(i, COL_���) = collist(i)("_type") & ""
                    .TextMatrix(i, COL_������) = collist(i)("_guarantor") & ""
                    .TextMatrix(i, COL_������) = collist(i)("_garnt_amount") & ""
                    .TextMatrix(i, COL_��ʱ����) = IIf(Nvl(collist(i)("_garnt_prop"), 0) = 1, "��", "")
                    .TextMatrix(i, COL_����ԭ��) = Nvl(collist(i)("_garnt_reason"))
                    .TextMatrix(i, COL_�Ǽ�ʱ��) = collist(i)("_create_time") & ""
                    .TextMatrix(i, COL_����ʱ��) = Nvl(collist(i)("_due_time"))
                    .TextMatrix(i, COL_ɾ����־) = collist(i)("_is_del") & ""
                    .TextMatrix(i, COL_����Ա����) = collist(i)("_operator_name") & ""
                    .TextMatrix(i, COL_����Ա���) = collist(i)("_operator_code") & ""
                    .TextMatrix(i, COL_ɾ������Ա����) = collist(i)("_del_operator_name") & ""
                    .TextMatrix(i, COL_ɾ������Ա���) = collist(i)("_del_operator_code") & ""
                    .TextMatrix(i, COL_ɾ��ʱ��) = collist(i)("_del_time") & ""
                Next
                .Redraw = True
            End With
        Else
            msh.Rows = 1 '����ձ�ͷ
            msh.Rows = 2
            msh.FixedRows = 1
        End If
        
        Call GetSuretyBalance
        Curdate = zldatabase.Currentdate
        For lngRow = 1 To msh.Rows - 1
            If UBound(Split(Trim(msh.TextMatrix(lngRow, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(lngRow, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            '����21368 by lesfeng 2010-08-02
            strɾ����־ = Trim(msh.TextMatrix(lngRow, GetColNum("ɾ����־")))
            
            If lng��ҳid = RowPageid And (Trim(msh.TextMatrix(lngRow, GetColNum("����ʱ��"))) = "" Or Trim(msh.TextMatrix(lngRow, GetColNum("����ʱ��"))) > Curdate) Then
                msh.Row = lngRow
                For i = 0 To msh.Cols - 1
                    msh.Col = i
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                        msh.CellForeColor = &HC00000
                    Else
                        msh.CellForeColor = &HFF&
                    End If
                Next
            Else
                 For i = 0 To msh.Cols - 1
                    msh.Col = i
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                    Else
                        msh.CellForeColor = &HFF&
                    End If
                Next
            End If
            
        Next lngRow
    End If
    msh.Row = 1
    msh.Col = 0: msh.ColSel = msh.Cols - 1
    Call msh_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPrepay(Optional ByVal blnClear As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ʷ��Ԥ������
    '����:������
    '����:2013-03-11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lng��ҳid As Long
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim collist As Collection
    Dim i As Long
    
    If mrsInfo Is Nothing Then
        lng����ID = mlng����ID
        lng��ҳid = mlng��ҳID
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = mlng����ID
        lng��ҳid = mlng��ҳID
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        lng��ҳid = Val(Nvl(mrsInfo!��ҳID))
    End If
    
    On Error GoTo errHandle
    
    If blnClear = True Then
        mshList.Clear
        mshList.Rows = 2
        Call SetDepositHeader
    Else
        '������ʷ�ɿ���ϸ�嵥
        strJsonIn = GetNode("pati_id", lng����ID, True)
        strJsonIn = strJsonIn & GetNode("pati_pageid", lng��ҳid)
        strJsonIn = strJsonIn & GetNode("type", 2)
        strJsonIn = GetNode("input", "{" & strJsonIn & "}", True, 1)
        
        If CallService("Zl_Exsesvr_Getdepositlist", strJsonIn, strJsonOut, "��ȡԤ����¼", P������Ժ����) Then
            Set collist = GetJsonListValue("output.item_list")
        Else
            Set collist = New Collection
        End If
        If collist.Count > 0 Then
            With mshList
                .Redraw = False
                .Rows = collist.Count + 1
                For i = 1 To collist.Count
                    .TextMatrix(i, COL_����) = collist(i)("_create_date")
                    .TextMatrix(i, COL_���ݺ�) = collist(i)("_bill_no")
                    .TextMatrix(i, COL_����) = collist(i)("_dept_name") & ""
                    .TextMatrix(i, COL_���) = collist(i)("_money")
                    .TextMatrix(i, COL_�ɿ���) = Format(collist(i)("_money"), "##,##0.00")
                    .TextMatrix(i, COL_����) = collist(i)("_blnc_mode")
                    .TextMatrix(i, COL_�տ���) = collist(i)("_operator_name")
                Next
                .Redraw = True
            End With
        Else
            mshList.Clear
            mshList.Rows = 2
        End If
         Call SetDepositHeader
    End If
     
    If mshList.Rows > 1 Then
        mshList.Row = 1: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check������Ϣ() As Boolean
    Check������Ϣ = True
    If mlngModul = P������Ժ���� Then
        If mrsInfo Is Nothing Then
            Check������Ϣ = False
        ElseIf mrsInfo.State = adStateClosed Then
            Check������Ϣ = False
        End If
        
        If Check������Ϣ = False Then
            stbThis.Panels(1).Text = "û��ȷ��Ҫ���е����Ĳ���!"
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    If Trim(txt������.Text) = "" Then
        stbThis.Panels(1).Text = "�����뵣��������,�����˲���Ϊ��!"
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If Not IsNumeric(txt������.Text) Then
        stbThis.Panels(1).Text = "��������ȷ�ĵ�����,������Ҫ������ֵ!"
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    ElseIf Val(txt������.Text) = 0 Then
        stbThis.Panels(1).Text = "�����뵣����,�������Ϊ��!"
        If txt������.Enabled Then txt������.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
    If chk��ʱ����.Value = 1 Then
        If Not IsNull(dtp����ʱ��.Value) Or chkUnlimit.Value = 1 Then
            stbThis.Panels(1).Text = "��ʱ�������������õ���ʱ�޻��޵�����!"
            If chk��ʱ����.Enabled Then chk��ʱ����.SetFocus
            Check������Ϣ = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        stbThis.Panels(1).Text = "����ԭ�������������� 25 �����ֻ� 50 ���ַ���"
        txtReason.SetFocus
        Check������Ϣ = False
        Exit Function
    End If
    
End Function

Private Sub cmdAdd_Click()
    Dim str������ As String, str����ʱ�� As String
    Dim strJsonIn As String, strJsonOut As String
    Dim i As Integer, Curdate As Date, blnδ���� As Boolean, bln��ʱ As Boolean, RowPageid As Integer
    Dim strɾ����־ As String
    
    '1.���ݼ��
    If Not Check������Ϣ Then Exit Sub
    
    Curdate = zldatabase.Currentdate
    
    For i = 1 To msh.Rows - 1 '�жϱ���סԺδ���ڵĵ�����¼��������ʾ
         If Trim(msh.TextMatrix(i, GetColNum("���"))) <> "" Then
            If UBound(Split(Trim(msh.TextMatrix(i, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(i, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            If mlng��ҳID = RowPageid Then
                '����21368 by lesfeng 2010-08-02
                strɾ����־ = Trim(msh.TextMatrix(i, GetColNum("ɾ����־")))
               If (Trim(Nvl(msh.TextMatrix(i, GetColNum("����ʱ��")))) = "" Or Nvl(msh.TextMatrix(i, GetColNum("����ʱ��"))) > Curdate) And strɾ����־ = "" Then
                   bln��ʱ = Nvl(msh.TextMatrix(i, GetColNum("��ʱ����"))) = "��"
                   blnδ���� = True: Exit For
               End If
            End If
        End If
    Next
    
    If blnδ���� Then
        If MsgBox("����δ���ڵ�" & IIf(bln��ʱ, "��ʱ", "") & "������¼����������" & IIf(bln��ʱ, "��֮ǰ����ʱ�����Զ�ʧЧ", "�ۼƵ���") & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    str������ = Replace(Trim(txt������.Text), "'", "''")
    str����ʱ�� = ""
    If Not IsNull(dtp����ʱ��.Value) Then str����ʱ�� = Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")
    
    '���ȼ��
    If Not CheckLen(txt������, 64) Then Exit Sub
    
    '2.���ݱ���
    On Error GoTo errH
    strJsonIn = GetJsonSurety(1, mlng����ID, mlng��ҳID, str������, Val(txt������.Text), chk��ʱ����.Value, Trim(txtReason.Text), str����ʱ��)
    If strJsonIn <> "" Then
        Call CallService("Zl_Exsesvr_Updatepatisurety", strJsonIn, strJsonOut, "����������¼", P������Ժ����, True, , , , True)
    End If
    
    '3.����ˢ��
    stbThis.Panels(1).Text = "������Ϣ�ѱ���!"
    Call LoadSurety
    Call Init������Ϣ
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    
    If cmdExit.Caption = "ȡ��(&C)" Then
        cmdModify.Caption = "�޸�(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "�˳�(&X)"
        fraEdit.Enabled = True      'SetCanEdit���ٴ�����
       
        'ˢ������,���ǲ�������
        stbThis.Panels(1).Text = ""
        Call LoadSurety
        Call Init������Ϣ
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim obj As Control
    If mlngModul = P������Ϣ���� Then
       If KeyCode = vbKeyEscape Then
            Call cmdExit_Click
        End If
    Else
        Select Case KeyCode
        Case vbKeyEscape
            Call cmdExit_Click
        Case vbKeyF2
            Call cmdPati_Click
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC����")
                If intIndex <= 0 Then Exit Sub
                 IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF11
            If txtPatient.Enabled And Not txtPatient.Locked Then txtPatient.SetFocus
        Case vbKeyReturn
            Set obj = Me.ActiveControl
            If InStr(1, ",txt������,txt������,dtp����ʱ��,chk��ʱ����,chkUnlimit,txtReason,", "," & obj.Name & ",") > 0 Then
               ' Call zlCommFun.PressKey(vbKeyTab)
            End If
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With tbcPage
        .Left = fraPati.Left
        .Top = fraPati.Top + fraPati.Height
        .Width = fraPati.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
    
    PicDeposit.Width = picSurety.Width
    PicDeposit.Height = picSurety.Height
    
    With msh
        .Width = picSurety.ScaleWidth
        .Height = picSurety.ScaleHeight - .Top
    End With
    
    With lblDeposit
        .Left = 60
        .Top = 60
    End With
    
    With mshList
        .Top = lblDeposit.Top + lblDeposit.Height + 60
        .Left = 0
        .Width = msh.Width
        .Height = PicDeposit.ScaleHeight - .Top
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    If cmdModify.Caption = "����(&S)" Then
        If MsgBox("��ǰ�޸ĵ���Ϣδ����,ȷʵҪ�˳���?", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Cancel = 1
    End If
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    
    If mlngModul = P������Ժ���� Then
        If Not mobjIDCard Is Nothing Then
            Call mobjIDCard.SetEnabled(False)
            Set mobjIDCard = Nothing
        End If
        If Not mobjICCard Is Nothing Then
            Call mobjICCard.SetEnabled(False)
            Set mobjICCard = Nothing
        End If
        Call zlCardSquareObject(True)
        Set mrsInfo = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlOneCardComLib.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
               Set mobjICCard = New clsICCard
               Call mobjICCard.SetParent(Me.hWnd)
               Set mobjICCard.gcnOracle = gcnOracle
        End If
           If Not mobjICCard Is Nothing Then
               txtPatient.Text = mobjICCard.Read_Card()
               If txtPatient.Text <> "" Then
                   Call txtPatient_KeyPress(vbKeyReturn)
               End If
           End If
           Exit Sub
    End If
     
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If mobjOneCardCom.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
    Call txtPatient_GotFocus
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlOneCardComLib.Card, objPatiInfor As zlOneCardComLib.clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("IC����")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strCardNo
        Call txtPatient_KeyPress(vbKeyReturn)
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("���֤��")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub msh_EnterCell()
    Dim str����ʱ�� As String
    Dim Datsys As Date, RowPageid As Integer
    Dim strɾ����־ As String
        
    If mlngModul = P������Ժ���� Then
        If Val(msh.RowData(msh.Row)) <= 0 Then
            stbThis.Panels(1).Text = ""
            cmdModify.Enabled = False
            cmdDel.Enabled = False
            Exit Sub
        End If
    End If
    
   '��ǰ����ҳ�벡����ҳ��ͬʱ�������޸�ɾ��,�ѹ��ڲ������޸�ɾ��
    Datsys = zldatabase.Currentdate
    
    '����21368 by lesfeng 2010-08-02
    strɾ����־ = Trim(msh.TextMatrix(msh.Row, GetColNum("ɾ����־")))
    
    If cmdModify.Caption = "�޸�(&M)" Then
        If mlng��ҳID = 0 And Trim(msh.TextMatrix(msh.Row, GetColNum("���"))) = "����" Then
            '����21368 by lesfeng 2010-08-02
            If strɾ����־ = "" Then
                cmdModify.Enabled = True
                cmdDel.Enabled = True
                stbThis.Panels(1).Text = "��ǰ������¼��Ч"
            Else
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
            End If
        Else
            If UBound(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("���"))), "��סԺ")) > 0 Then 'ȡ��ѡ������ҳID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("���"))), "��סԺ")(0), "��")(1))
            Else
                RowPageid = 0
            End If
            If mlng��ҳID <> RowPageid Then
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "��ǰ������¼�Ǳ���סԺ������"
            Else
                str����ʱ�� = Trim(msh.TextMatrix(msh.Row, GetColNum("����ʱ��")))
            
                If str����ʱ�� <> "" Then
                    If CDate(str����ʱ��) < Datsys Then
                         cmdModify.Enabled = False
                         cmdDel.Enabled = False
                        '����21368 by lesfeng 2010-08-02
                         If strɾ����־ = "" Then
                            stbThis.Panels(1).Text = "��ǰ������¼�ѹ���"
                        Else
                            stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                        End If
                    Else
                        '����21368 by lesfeng 2010-08-02
                        If strɾ����־ = "" Then
                            cmdModify.Enabled = True
                            cmdDel.Enabled = True
                            stbThis.Panels(1).Text = "��ǰ������¼��Ч"
                        Else
                            cmdModify.Enabled = False
                            cmdDel.Enabled = False
                            stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                        End If
                    End If
                Else
                    '����21368 by lesfeng 2010-08-02
                    If strɾ����־ = "" Then
                        cmdModify.Enabled = True
                        cmdDel.Enabled = True
                        stbThis.Panels(1).Text = "��ǰ������¼��Ч"
                    Else
                        cmdModify.Enabled = False
                        cmdDel.Enabled = False
                        stbThis.Panels(1).Text = "��ǰ������¼�Ѿ����ɾ��"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim dblMoney As Double, lngRow As Long
    
    If txtPatient.Locked Then Exit Sub
        
    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    If txtPatient.Tag <> "" Then Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        If frmPatiSelect.ShowMe(Me) = False Then
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
        End If
    End If
    Me.Refresh
    mstr�������� = ""
    txtPatient.ForeColor = &HFF0000
    
    'ˢ����ϻ���������س�
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '��ȡ������Ϣ
        Call ClearWinInfor(True)
        
        If IDKind.GetCurCard.���� Like "IC��*" And IDKind.GetCurCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        If Not GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCancel, blnCard) Then
            If blnCancel Then 'ȡ������
                Call zlControl.TxtSelAll(txtPatient): txtPatient.SetFocus: Exit Sub
            End If
            stbThis.Panels(1).Text = "δ�ҵ��ò��ˣ�������������!"
            If blnCard = True Then
                txtPatient.PasswordChar = "": txtPatient.Text = "": txtPatient.IMEMode = 0
            Else
                txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
            End If
            Set mrsInfo = New ADODB.Recordset
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Else
            '���ò��˷�����Ϣ
            mlng����ID = Val(Nvl(mrsInfo!����ID, 0))
            mlng��ҳID = Val(Nvl(mrsInfo!��ҳID, 0))
            
            Call ClearWinInfor
            If mrsInfo!��ǰ����ID <> 0 Then
                lblBed.Caption = "���ţ�" & IIf(mrsInfo!���� = 0, "��ͥ", mrsInfo!����)
            End If
            
            lblNO.Caption = "סԺ�ţ�" & IIf(mrsInfo!סԺ�� = 0, "", mrsInfo!סԺ��)
            lblDept.Caption = "���ң�" & GET��������(mrsInfo!��ǰ����ID)
            
            lblType.Caption = "�ѱ�ȼ���" & mrsInfo!�ѱ�
'            lbl������.Caption = lbl������.Tag & mrsInfo!������
'            lbl�������.Caption = lbl�������.Tag & mrsInfo!������
'            chk����temp.Value = mrsInfo!��������
            
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            txtPatient.Text = mrsInfo!����
            txtPatient.Tag = mrsInfo!����ID
            '-----------------------------------------------------------------------------------------
            lblSex.Caption = "�Ա�" & IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�)
            lblAge.Caption = "���䣺" & IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
'            lbl��ͥ��ַ.Caption = lbl��ͥ��ַ.Tag & Nvl(mrsInfo!��ͥ��ַ)
            lblCur.Caption = "ҽ�Ƹ��ʽ��" & Nvl(mrsInfo!ҽ�Ƹ��ʽ)
            dblMoney = 0
            For lngRow = 1 To mshList.Rows - 1
                 dblMoney = Format(dblMoney + Val(mshList.TextMatrix(lngRow, GetColNumList("���"))), "#0.00;-#0.00;0.00")
            Next
            lblDeposit.Caption = "Ԥ���ܶ" & IIf(dblMoney = 0, "", dblMoney)
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If mrsInfo Is Nothing Then
            cmdAdd.Enabled = False
        ElseIf mrsInfo.State = adStateClosed Then
            cmdAdd.Enabled = False
        Else
            cmdAdd.Enabled = True
        End If
    End If
End Sub

Private Function GetPatient(ByVal objCard As zlOneCardComLib.Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '���ܣ���ȡ������Ϣ
    '������strInput=[ˢ��]|[A����ID]|[BסԺ��]
    '˵����
    '     �Զ�ʶ������Ժ״̬,����(����ID,��ҳID,����,�Ա�,����,סԺ��,����,��Ժ��־)
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng�����ID As Long, bln�����ʻ� As Boolean, lng����ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    Dim colPati As Collection
    Dim colPage As Collection
    Dim colItem As Collection
    Dim colOtherFind As Collection
    Dim strFields As String
    Dim strKeyName As String
    Dim strKeyValue As String
    
    blnCancel = False
    strWhere = ""
      
    If (blnCard And objCard.���� Like "����*") _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   'ˢ����ȱʡ�Ŀ�
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If mobjOneCardCom.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '����ID
        lng����ID = Val(Mid(strInput, 2))
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  'סԺ��(��ס(��)Ժ�Ĳ���)
        strKeyName = "סԺ��": strKeyValue = Mid(strInput, 2)
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
        strKeyName = "�����": strKeyValue = Mid(strInput, 2)
    Else '��������
        Select Case objCard.����
            Case "����"
                If Not mblnSeekName Then
                    MsgBox "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��", vbInformation, gstrSysName
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                Else
                    If zlCommFun.IsCharChinese(strInput) Then
                        If Len(strInput) < 2 Then
                            MsgBox "��������""" & strInput & """���Ȳ���2λ��", vbInformation, gstrSysName
                            txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                        End If
                    Else
                        If Len(strInput) < 4 Then
                            MsgBox "��������""" & strInput & """���Ȳ���4λ��", vbInformation, gstrSysName
                            txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                        End If
                    End If
                    On Error GoTo errH
                    Set colOtherFind = New Collection
                    colOtherFind.Add Array("����", strInput & "%")
                    colOtherFind.Add Array("��ѯסԺ״̬", 1)
                    If zl_PatiSvr_GetPatiInfo(0, colOtherFind, colPati, 2) Then
                        strSQL = ""
                        For Each colItem In colPati
                            If strSQL <> "" Then strSQL = strSQL & "Union All " & vbNewLine
                            strSQL = strSQL & "Select " & colItem("_pati_id") & " As ����ID," & _
                                    Nvl(colItem("_pati_pageid"), "-NULL") & " As ��ҳID," & _
                                    "'" & colItem("_pati_name") & "' As ����," & _
                                    "'" & colItem("_pati_sex") & "' As �Ա�," & _
                                    "'" & colItem("_pati_age") & "' As ����," & _
                                    Nvl(colItem("_inpatient_num"), "-NULL") & " As סԺ��," & _
                                    "'" & colItem("_pati_dept_name") & "' As ����," & _
                                    "'" & colItem("_pati_birthdate") & "' As ��������," & _
                                    "'" & colItem("_pati_idcard") & "' As ���֤��," & _
                                    "'" & colItem("_pat_home_addr") & "' As ��ͥ��ַ," & _
                                    "'" & colItem("_adta_time") & "' As ��Ժʱ�� From Dual "
                        Next
                    End If
                    If strSQL <> "" Then
                        strPati = "Select A.����ID As ID,A.* From (" & strSQL & ") A Order by A.��Ժʱ�� DESC,A.����"
                        vRect = zlControl.GetControlRect(txtPatient.hWnd)
                        Set rsTmp = zldatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True)
                    End If
                    If Not rsTmp Is Nothing Then
                        lng����ID = rsTmp!����ID
                    Else
                        Set mrsInfo = New ADODB.Recordset: Exit Function
                    End If
                End If
            Case "ҽ����"
                strKeyName = "ҽ����": strKeyValue = UCase(strInput)
            Case "IC����"
                strInput = UCase(strInput)
                If mobjOneCardCom.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0) And objCard.ϵͳ
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                 strKeyName = "�����": strKeyValue = strInput
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strKeyName = "סԺ��": strKeyValue = strInput
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    bln�����ʻ� = objCard.�Ƿ�����ʻ�
                    If mobjOneCardCom.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If mobjOneCardCom.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                blnHavePassWord = True
        End Select
    End If
    
    If strKeyName <> "" And strKeyValue <> "" Then
        Set colOtherFind = New Collection
        colOtherFind.Add Array(strKeyName, strKeyValue)
        If Not zl_PatiSvr_GetPatiID(Nothing, colOtherFind, colPati) Then GoTo NotFoundPati:
        lng����ID = colPati(1)("_pati_id")
    End If
    
    If lng����ID <= 0 Then GoTo NotFoundPati:
    '��ѯ������Ϣ
    If Not zl_PatiSvr_GetPatiInfo(lng����ID, Nothing, colPati, 2) Then GoTo NotFoundPati:
    Set colPati = colPati(1)
             
    '��ѯ������ҳ
    strSQL = "{""input"":{""query_type"":1,""pati_pageids"":""" & lng����ID & ":" & colPati("_pati_pageid") & """,""is_lastpage"":1,""is_babyinfo"":0,""is_transdeptinfo"":1,""is_ex"":1,""is_bed"":0}}"
    Call CallService("zl_CIsSvr_GetPatiPageInfo", strSQL, , Me.Caption, , False, , , , True)
    Set colPage = gobjService.GetJsonListValue("output.page_list[0]")
    If colPage Is Nothing Or colPati Is Nothing Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    
    strFields = "����ID|adBigInt|18,��ҳID|adBigInt|18,����֤��||50,����||100,�Ա�||4,����||20,��������||50," & _
            "��ǰ����ID|adBigInt|18,����||10,סԺ��|adBigInt|18," & _
            "�ѱ�||10,������||100,������||20,��������|adSingle|1,ҽ�Ƹ��ʽ||20"
    Set mrsInfo = InitRS(strFields)
    With mrsInfo
        .AddNew
        !����ID = colPati("_pati_id"): !��ҳID = colPati("_pati_pageid"): !����֤�� = colPati("_card_captcha")
        !���� = colPati("_pati_name"): !�Ա� = colPati("_pati_sex"): !���� = colPati("_pati_age")
        !�������� = colPati("_pati_type"): !��ǰ����ID = colPage("_pati_dept_id"): !���� = Nvl(colPage("_pati_bed"))
        !סԺ�� = Nvl(colPage("_inpatient_num")): !�ѱ� = colPage("_fee_category")
        !ҽ�Ƹ��ʽ = colPage("_mdlpay_mode_name")
    End With
    '��ȡ������Ϣ
    strSQL = "{""input"":{""pati_pageid"":0,""pati_id"":" & lng����ID & ",""surety_prop"":1}}"
    Call CallService("Zl_Exsesvr_Getpatisurety", strSQL, , , , False, , , , True)
    If Val(gobjService.GetJsonNodeValue("output.guarantee_money") & "") <> 0 Then
        mrsInfo!�������� = Val(gobjService.GetJsonNodeValue("output.surety_prop") & "")
        mrsInfo!������ = gobjService.GetJsonNodeValue("output.entsurety") & ""
        mrsInfo!������ = Format(Val(gobjService.GetJsonNodeValue("output.guarantee_money") & ""), "0.00")
    End If
    mrsInfo.Update
    '��Ҫ��������
    If mblnCheckPass And (blnCard Or blnICCard) Then
        If Not blnHavePassWord Then
            strPassWord = Nvl(mrsInfo!����֤��)
        End If
        If strPassWord <> "" Then
            If CreatePublicExpense() Then
                If gobjPublicExpense.zlVerifyPassWord(Me, strPassWord, mrsInfo!����, mrsInfo!�Ա�, mrsInfo!����) = False Then
                     Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            End If
        End If
    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function


Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(True)
    txtPatient.Tag = ""
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    '����27379 by lesfeng 2010-01-18
    If mrsInfo.State = 1 Then
        mstr�������� = IIf(IsNull(mrsInfo!��������), "", mrsInfo!��������)
    End If
    If mstr�������� = "" Then
        If mrsInfo.State = 1 Then
            If GetOutPatient(mrsInfo!����ID) Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        txtPatient.ForeColor = ReadPatiColor(mstr��������, True)
    End If
End Sub

Private Function GetOutPatient(ByVal lngID As Long) As Boolean
'���ܣ��ж����ﲡ���Ƿ�����ҽ��
    Dim int���� As Integer
    Dim colPati As Collection
    
    GetOutPatient = False
    On Error GoTo errH
 
    If zl_PatiSvr_GetPatiInfo(lngID, Nothing, colPati) Then
        int���� = Val(colPati(1)("_insurance_type") & "")
        GetOutPatient = int���� <> 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '����27554 by lesfeng 2010-01-19 lngTXTProc �޸�ΪglngTXTProc
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtReason_GotFocus()
    zlControl.TxtSelAll txtReason
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txtReason, KeyAscii
    End If
End Sub

Private Sub txtReason_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OS.OpenImeByName
End Sub

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chkUnlimit.TabStop = (txt������.Text = "")
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc(".") And InStr(txt������.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt������_LostFocus()
    If IsNumeric(txt������.Text) Then
        If txt������.Text = "999999999" Then
            stbThis.Panels(1).Text = "�����������ֵ����ֵ��ʾ���޵�����"
            If txt������.Enabled Then txt������.SetFocus
        Else
            txt������.Text = Format(txt������.Text, "0.00")
        End If
    Else
        txt������.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txt������_GotFocus()
    zlControl.TxtSelAll txt������
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt������, KeyAscii
    End If
End Sub

Private Sub txt������_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OS.OpenImeByName
End Sub

Private Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
   
    'ֻ��:ִ�л��˷�ʱ,�ſ��ܹܽ��㿨��
    If blnClosed Then
       If Not mobjOneCardCom Is Nothing Then
            Call mobjOneCardCom.CloseWindows
            Set mobjOneCardCom = Nothing
        End If
        Exit Sub
    End If
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    Set mobjOneCardCom = CreateObject("zlOneCardComLib.clsOneCardComLib")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If mobjOneCardCom.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub

Private Function GetJsonSurety(ByVal bytFunc As Byte, ByVal lngPatiid As Long, ByVal lngPageID As Long, Optional ByVal strSurety As String, _
   Optional ByVal dblAmount As Double, Optional ByVal bytType As Byte, Optional ByVal strReason As String, Optional ByVal strDueTime As String, _
   Optional ByVal strCreateTime As String) As String
'����:��ȡ����JSON��
'����:
'      bytFunc          ����ID 1-����;2-����;3-ɾ��
'      lngPatiId        ����id
'      lngPageId        ��ҳID
'      strSurety        ������
'      dblAmount        ������
'      bytType          ��������
'      strReason        ����ԭ��
'      strDueTime       ����ʱ��   ��ʽ�� "yyyy-MM-dd HH:mm:ss"
'      strCreateTime    �Ǽ�ʱ��   ���»�ɾ��ʱ����
    Dim strIn As String
    
    strIn = GetNode("func_id", bytFunc, True)
    strIn = strIn & GetNode("pati_id", lngPatiid)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    If bytFunc <> 3 Then
        strIn = strIn & GetNode("guarantor", strSurety)
        strIn = strIn & GetNode("garnt_amount", dblAmount)
        strIn = strIn & GetNode("garnt_prop", bytType)
        strIn = strIn & GetNode("garnt_reason", strReason)
        If strDueTime <> "" Then
            strIn = strIn & GetNode("due_time", strDueTime)
        End If
    End If

    strIn = strIn & GetNode("operator_code", UserInfo.���)
    strIn = strIn & GetNode("operator_name", UserInfo.����)

    If bytFunc <> 1 Then
        strIn = strIn & GetNode("create_time", strCreateTime)
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    GetJsonSurety = strIn
End Function

