VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmPatiFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   6480
      TabIndex        =   5
      Top             =   2475
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   4
      Top             =   735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6480
      TabIndex        =   3
      Top             =   300
      Width           =   1100
   End
   Begin VB.Frame fraBdr 
      Height          =   3420
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6225
      Begin VB.ComboBox cboPayPlan 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1800
         Width           =   2085
      End
      Begin VB.CheckBox chk�Ǽ� 
         Caption         =   "�Ǽ�ʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   338
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "��������"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   713
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk��Ժ 
         Caption         =   "��Ժʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2610
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   255
         Left            =   5730
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   1450
         Width           =   285
      End
      Begin VB.CheckBox chk��Ժ 
         Caption         =   "��Ժʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   2985
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   3945
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3945
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1425
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtp��ԺE 
         Height          =   300
         Left            =   3945
         TabIndex        =   15
         Top             =   2925
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺB 
         Height          =   300
         Left            =   1230
         TabIndex        =   16
         Top             =   2925
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺE 
         Height          =   300
         Left            =   3945
         TabIndex        =   20
         Top             =   2550
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺB 
         Height          =   300
         Left            =   1230
         TabIndex        =   21
         Top             =   2550
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp����E 
         Height          =   300
         Left            =   3945
         TabIndex        =   24
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp�Ǽ�E 
         Height          =   300
         Left            =   3945
         TabIndex        =   27
         Top             =   285
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp����B 
         Height          =   300
         Left            =   1230
         TabIndex        =   31
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp�Ǽ�B 
         Height          =   300
         Left            =   1230
         TabIndex        =   32
         Top             =   285
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1230
         TabIndex        =   30
         Top             =   1425
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1425
         Width           =   2085
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1230
         MaxLength       =   18
         TabIndex        =   29
         Top             =   1050
         Width           =   2085
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   1230
         TabIndex        =   19
         Top             =   2175
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPatiFilter.frx":0000
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
         IDKindAppearance=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         IDKindWidth     =   555
         FindPatiShowName=   0   'False
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblҽ�Ƹ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ��ʽ"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label lbl�Ǽ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   28
         Top             =   345
         Width           =   180
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   25
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lbl��Ժ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   22
         Top             =   2610
         Width           =   180
      End
      Begin VB.Label lbl��Ժ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   17
         Top             =   2985
         Width           =   180
      End
      Begin VB.Label lblIDKind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   840
         TabIndex        =   12
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   630
         TabIndex        =   11
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3450
         TabIndex        =   9
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label lbl�ѱ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   810
         TabIndex        =   8
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   3450
         TabIndex        =   7
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   1485
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPatiFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mbytType As Byte '��:�����嵥����0-����,1-��Ժ,2-��Ժ,3-����,4-����
Public mstrFilter As String '��:����
Public mstrFilterInfo As String '������Ϣ ר�ù�������
Public mbytInFun As Byte '0-��ͨ����,1-���ⲡ�˹��˵���
Public mlngPatiId As Long   '����ID

Private mstrFindType As String        '�洢��ǰ������������

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        zlControl.TxtSelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub chk�Ǽ�_Click()
    If chk�Ǽ�.Tag <> "" Then chk�Ǽ�.Value = 0: Exit Sub
    dtp�Ǽ�B.Enabled = (chk�Ǽ�.Value = 1)
    dtp�Ǽ�E.Enabled = dtp�Ǽ�B.Enabled
    If dtp�Ǽ�B.Enabled Then dtp�Ǽ�B.SetFocus
End Sub

Private Sub chk����_Click()
    If chk����.Tag <> "" Then chk����.Value = 0: Exit Sub
    dtp����B.Enabled = (chk����.Value = 1)
    dtp����E.Enabled = dtp����B.Enabled
    If dtp����B.Enabled Then dtp����B.SetFocus
End Sub

Private Sub chk��Ժ_Click()
    If chk��Ժ.Tag <> "" Then chk��Ժ.Value = 0: Exit Sub
    dtp��ԺB.Enabled = (chk��Ժ.Value = 1)
    dtp��ԺE.Enabled = dtp��ԺB.Enabled
    If dtp��ԺB.Enabled Then dtp��ԺB.SetFocus
End Sub

Private Sub chk��Ժ_Click()
    If chk��Ժ.Tag <> "" Then chk��Ժ.Value = 0: Exit Sub
    dtp��ԺB.Enabled = (chk��Ժ.Value = 1)
    dtp��ԺE.Enabled = dtp��ԺB.Enabled
    If dtp��ԺB.Enabled Then dtp��ԺB.SetFocus
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    txtסԺ��.Text = Trim(txtסԺ��.Text)
     
    If txtסԺ��.Text = "" And Trim(PatiIdentify.Text) = "" Then
        If chk�Ǽ�.Value = 0 And chk��Ժ.Value = 0 And chk��Ժ.Value = 0 And mbytType <> 1 Then
            MsgBox "������ѡ��һ���Ǽ�ʱ�䷶Χ.", vbInformation, gstrSysName
            chk�Ǽ�.Value = 1
            Exit Sub
        End If
        
        If mbytType = 0 Then
            If chk�Ǽ�.Value = 0 Then
                MsgBox "������ѡ��һ���Ǽ�ʱ�䷶Χ.", vbInformation, gstrSysName
                chk�Ǽ�.Value = 1
                Exit Sub
            End If
        End If
    End If
        
    Call MakeFilter
    gblnOK = True
    Hide
End Sub

Private Sub dtp����E_Change()
    dtp����B.MaxDate = dtp����E.Value
End Sub

Private Sub dtp��ԺE_Change()
    dtp��ԺB.MaxDate = dtp��ԺE.Value
End Sub

Private Sub dtp�Ǽ�E_Change()
    dtp�Ǽ�B.MaxDate = dtp�Ǽ�E.Value
End Sub

Private Sub dtp��ԺE_Change()
    dtp��ԺB.MaxDate = dtp��ԺE.Value
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Select Case mbytType
        Case 0
            dtp�Ǽ�B.SetFocus
        Case 1
            chk��Ժ.SetFocus
        Case 2
            dtp��ԺB.SetFocus
        Case 3, 4
            dtp�Ǽ�B.SetFocus
    End Select
    mlngPatiId = 0
    PatiIdentify.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, datTmp As Date, i As Integer
    Dim strKindType As String
    
    If Not gobjSquare Is Nothing Then
        strKindType = "��|����|0|0|0|0|0|0|0|0|0;��|���￨|0|0|0|0|0|0|0|0|0;��|�����|0|0|0|0|0|0|0|0|0;ҽ|ҽ����|0|0|0|0|0|0|0|0|0;��|�������֤|0|0|0|0|0|0|0|0|0;IC|IC��|0|0|0|0|0|0|0|0|0;��|�ֻ���|0|0|0|0|0|0|0|0|0"
        Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKindType, "ZL9Patient")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
    End If
    PatiIdentify.Text = ""
    lbl�ѱ�.Visible = mbytInFun = 0
    cbo�ѱ�.Visible = mbytInFun = 0
    lbl���.Visible = mbytInFun = 1
    txt���.Visible = mbytInFun = 1
    
    If mbytInFun = 0 Then
        '�ѱ�
        If glngSys Like "8??" Then
            lbl�ѱ�.Caption = "��Ա�ȼ�"
        Else
            If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
                lbl�ѱ�.Caption = "����ѱ�"
            Else
                lbl�ѱ�.Caption = "סԺ�ѱ�"
            End If
        End If
        
        Set rsTmp = Nothing
        Set rsTmp = GetDictData("�ѱ�")
        cbo�ѱ�.Clear
        cbo�ѱ�.AddItem "���зѱ�"
        cbo�ѱ�.ListIndex = 0
        If Not rsTmp Is Nothing Then
            For i = 1 To rsTmp.RecordCount
                cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
                rsTmp.MoveNext
            Next
        End If
    ElseIf mbytInFun = 1 Then
        chk�Ǽ�.Caption = "����ʱ��"
    End If
    
    '�Ա�
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("�Ա�")
    cbo�Ա�.Clear
    cbo�Ա�.AddItem "�����Ա�"
    cbo�Ա�.ListIndex = 0
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    'ҽ�Ƹ��ʽ
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("ҽ�Ƹ��ʽ")
    cboPayPlan.Clear
    cboPayPlan.AddItem "���з�ʽ"
    cboPayPlan.ListIndex = 0
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cboPayPlan.AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If


    '���ó�ʼ����
    On Error Resume Next    '����ע���洢��Чʱ��ʱ����
    Curdate = zlDatabase.Currentdate
    dtp�Ǽ�B.MaxDate = Format(DateAdd("d", 1, Curdate), dtp�Ǽ�E.CustomFormat)
    dtp����B.MaxDate = Format(Curdate, dtp����E.CustomFormat)
    dtp��ԺB.MaxDate = Format(DateAdd("d", 1, Curdate), dtp��ԺE.CustomFormat)
    dtp��ԺB.MaxDate = Format(DateAdd("d", 1, Curdate), dtp��ԺE.CustomFormat)
        
    datTmp = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtp�Ǽ�B.Value = datTmp
    datTmp = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtp�Ǽ�E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "������ʼʱ��", Format(DateAdd("yyyy", -100, Curdate), "yyyy-MM-dd")))
    dtp����B.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��������ʱ��", Format(dtp����B.MaxDate, dtp����E.CustomFormat)))
    dtp����E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(Curdate, "YYYY-MM-DD")))
    dtp��ԺB.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(dtp��ԺB.MaxDate, dtp��ԺE.CustomFormat)))
    dtp��ԺE.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(Curdate, "YYYY-MM-DD")))
    dtp��ԺB.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(dtp��ԺB.MaxDate, dtp��ԺE.CustomFormat)))
    dtp��ԺE.Value = datTmp
    
    On Error GoTo 0
    
    
    Select Case mbytType
        Case 0 '���в���
            chk�Ǽ�.Value = 1
            chk����.Value = 0
            chk��Ժ.Value = 0
            chk��Ժ.Value = 0
        Case 1 '��Ժ����
            chk�Ǽ�.Value = 0
            chk����.Value = 0
            chk��Ժ.Value = 0
            chk��Ժ.Value = 0: chk��Ժ.Tag = 1
        Case 2 '��Ժ����
            chk�Ǽ�.Value = 0
            chk����.Value = 0
            chk��Ժ.Value = 0
            chk��Ժ.Value = 1
        Case 3, 4 '���ﲡ��
            chk�Ǽ�.Value = 1
            chk����.Value = 0
            chk��Ժ.Value = 0: chk��Ժ.Tag = 1
            chk��Ժ.Value = 0: chk��Ժ.Tag = 1
    End Select
    
    If glngSys Like "8??" And Not Visible Then
        chk��Ժ.Visible = False
        dtp��ԺB.Visible = False
        dtp��ԺE.Visible = False
        lbl��Ժ.Visible = False
        chk��Ժ.Visible = False
        dtp��ԺB.Visible = False
        dtp��ԺE.Visible = False
        lbl��Ժ.Visible = False
        fraBdr.Height = fraBdr.Height - 900
        Me.Height = Me.Height - 900
        cmdOK.Top = cmdOK.Top - 100
        cmdCancel.Top = cmdCancel.Top - 100
        cmdDef.Top = cmdDef.Top - 800
    End If
End Sub

Public Sub MakeFilter()
    
    mstrFilter = ""
    mstrFilterInfo = "" 'ֻ�����������е�����
    If chk�Ǽ�.Value = 1 Then
        mstrFilter = mstrFilter & " And A.�Ǽ�ʱ�� Between [3] And [4]"
        mstrFilterInfo = mstrFilterInfo & " And A.�Ǽ�ʱ�� Between [3] And [4]"
    End If
    If chk����.Value = 1 Then mstrFilter = mstrFilter & " And A.�������� Between [5] And [6]"
    If chk��Ժ.Value = 1 Then mstrFilter = mstrFilter & " And P.��Ժ���� Between [7] And [8]"
    If chk��Ժ.Value = 1 Then mstrFilter = mstrFilter & " And P.��Ժ���� Between [9] And [10]"
    
    If txtסԺ��.Text <> "" Then
        mstrFilter = mstrFilter & " And A.����ID = (Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[11])"
        mstrFilterInfo = mstrFilterInfo & " And A.����ID = (Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[11])"
    End If
    If cbo�Ա�.ListIndex <> 0 Then mstrFilter = mstrFilter & " And A.�Ա�=[12]"
    If Trim(txt����.Text) <> "" Then mstrFilter = mstrFilter & " And A.����=[13]"
    
    '���������������ⲡ�˹���
    If txt���.Visible Then
        If txt���.Text <> "" Then mstrFilter = mstrFilter & " And C.���=[14]"
    Else
        '��ͬ�Ĳ鿴��Χʱ������ͬ
        If cbo�ѱ�.ListIndex <> 0 Then
            If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
                mstrFilter = mstrFilter & " And A.�ѱ�=[14]"
            Else
                mstrFilter = mstrFilter & " And P.�ѱ�=[14]"
            End If
        End If
    End If

    If Trim(PatiIdentify.Text) <> "" Then
        If mlngPatiId = 0 Then
            Select Case PatiIdentify.GetCurCard.����  '"1-����;2-���￨;3-�����;4-ҽ����;5-���֤��;6-IC����;7-�ֻ���"
                Case "����"
                    If chk�Ǽ�.Value = 1 Or chk��Ժ.Value = 1 Or chk��Ժ.Value = 1 Then
                        mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���� like [15]"
                        mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���� like [15]"
                    Else
                        mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����=[15]"
                        mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����=[15]"
                    End If
                Case "���￨"
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���￨��=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���￨��=[15]"
                Case "�����"
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�����=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�����=[15]"
                Case "ҽ����"
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.ҽ����=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.ҽ����=[15]"
                Case "�������֤"
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���֤��=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���֤��=[15]"
                Case "IC��"
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.IC����=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.IC����=[15]"
                Case "�ֻ���"
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�ֻ���=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�ֻ���=[15]"
            End Select
        Else
            mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����ID=[15]"
            mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����ID=[15]"
        End If
    End If
    If cboPayPlan.ListIndex <> 0 Then
        If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
            mstrFilter = mstrFilter & " And A.ҽ�Ƹ��ʽ=[17]"
        Else
            mstrFilter = mstrFilter & " And P.ҽ�Ƹ��ʽ=[17]"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0

    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "������ʼʱ��", Format(Me.dtp����B.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��������ʱ��", Format(Me.dtp����E.Value, "yyyy-MM-dd 23:59:59")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(Me.dtp��ԺB.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(Me.dtp��ԺE.Value, "yyyy-MM-dd 23:59:59")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(Me.dtp��ԺB.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(Me.dtp��ԺE.Value, "yyyy-MM-dd 23:59:59")
End Sub

Private Sub PatiIdentify_Change()
    mlngPatiId = 0
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    mlngPatiId = 0
    If objHisPati Is Nothing Then Exit Sub
    mlngPatiId = objHisPati.����ID
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    If mstrFindType = objCard.���� And InStr(";���￨;�����;����;�������֤;IC��;ҽ����;�ֻ���;", ";" & mstrFindType & ";") > 0 Then
        'ѡ�����Ͳ�ͬ
        mlngPatiId = 0
        blnCancel = True
    End If
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mstrFindType = objCard.����
    PatiIdentify.Text = ""
    mlngPatiId = 0
End Sub

Private Sub txt���_GotFocus()
    Call zlControl.TxtSelAll(txt���)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    Call OpenIme(gstrIme)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt����
                txt����.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txtסԺ��_GotFocus()
    Call zlControl.TxtSelAll(txtסԺ��)
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

