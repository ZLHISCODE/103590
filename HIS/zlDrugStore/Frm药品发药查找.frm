VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form FrmҩƷ��ҩ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2970
   ClientLeft      =   3255
   ClientTop       =   4680
   ClientWidth     =   6795
   Icon            =   "FrmҩƷ��ҩ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5460
      TabIndex        =   25
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4230
      TabIndex        =   24
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame fra�������� 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   1155
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   6555
      Begin VB.ComboBox Cbo��ҩ�� 
         Height          =   300
         Left            =   4230
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2085
      End
      Begin VB.ComboBox Cbo������ 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   270
         Width           =   2085
      End
      Begin VB.CommandButton CmdҩƷ 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   23
         Top             =   660
         Width           =   285
      End
      Begin VB.TextBox TxtҩƷ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   22
         Top             =   660
         Width           =   4725
      End
      Begin VB.CheckBox ChkҩƷ 
         Caption         =   "ҩƷ(&P)"
         Height          =   210
         Left            =   270
         TabIndex        =   21
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Lbl��ҩ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   19
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   17
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame Fra�������� 
      Caption         =   "��������"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6555
      Begin VB.CheckBox chkSend 
         Caption         =   "��Ժ��ҩ"
         Height          =   180
         Index           =   1
         Left            =   5280
         TabIndex        =   29
         Top             =   1860
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkSend 
         Caption         =   "Ժ����ҩ"
         Height          =   180
         Index           =   0
         Left            =   4200
         TabIndex        =   28
         Top             =   1860
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtҽ���� 
         Height          =   300
         Left            =   930
         TabIndex        =   26
         Top             =   1800
         Width           =   2085
      End
      Begin VB.TextBox txt���￨ 
         Height          =   300
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         Left            =   930
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1440
         Width           =   2085
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   930
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1050
         Width           =   2085
      End
      Begin VB.ComboBox Cbo���� 
         Height          =   276
         Left            =   4200
         TabIndex        =   15
         Text            =   "Cbo����"
         Top             =   1440
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker Dtp��ʼDate 
         Height          =   300
         Left            =   930
         TabIndex        =   2
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   85524483
         CurrentDate     =   37007
      End
      Begin VB.TextBox Txt����NO 
         Height          =   300
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   8
         Top             =   660
         Width           =   2085
      End
      Begin VB.TextBox Txt��ʼNO 
         Height          =   300
         Left            =   930
         MaxLength       =   8
         TabIndex        =   6
         Top             =   660
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker Dtp����Date 
         Height          =   300
         Left            =   4200
         TabIndex        =   4
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   85524483
         CurrentDate     =   37007
      End
      Begin zlIDKind.IDKindNew IDKNType 
         Height          =   300
         Left            =   3240
         TabIndex        =   31
         Top             =   1050
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         ShowSortName    =   0   'False
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
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl��ҩ���� 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ����"
         Height          =   180
         Left            =   3360
         TabIndex        =   30
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label lblҽ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   27
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label lblסԺ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   12
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   9
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3780
         TabIndex        =   14
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label Lbl����Date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3420
         TabIndex        =   3
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Lbl��ʼDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Lbl����NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����NO"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   7
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Lbl��ʼNO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼNO"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   720
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmҩƷ��ҩ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--����߶�--
Private Const DblNormalHeight As Double = 3330
Private Const DblAdvanceHeight As Double = 4845
Private FrmObj As Form
Private mstrPrivs As String                             'Ȩ�ޣ���������Ƿ��и�������ѡ���Ȩ�ޣ��Ծ�������Ĵ�С

'--������ʹ��--
Private BlnStartUp As Boolean                           '�����ɹ�
Private strReturn As String
Private BlnState As Boolean                             '״̬(��״̬��������߶ȼ������SQL)
Private mbln���￨ As Boolean

Private mobjSquareCard As Object             'һ��ͨ�ӿ�

Private mlng����ID As Long

'--�ⲿ�������--
Private lngҩ��ID As Long                               '�ⷿID
Private Int���� As Integer                              '����
Private IntOper As Integer

Private Type Type_SQLCondition
    date��ʼ���� As Date
    date�������� As Date
    str��ʼNO As String
    str����NO As String
    str���� As String
    str���￨ As String
    str��ʶ�� As String
    lng����ID As Long
    str������ As String
    str����� As String
    lngҩƷid As Long
    strҽ���� As String
End Type

Private SQLCondition As Type_SQLCondition

Private mint��Ժ��ҩ As Integer

Private Sub Cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "A,D"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo����.ListCount = 0 Then Exit Sub
    
    If Cbo����.ListIndex >= 0 Then
        If Val(Cbo����.Tag) = Cbo����.ItemData(Cbo����.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, Cbo����, Trim(Cbo����.Text), str��������, , "1,2,3") = False Then
        Exit Sub
    End If
    If Cbo����.ListIndex >= 0 Then
        Cbo����.Tag = Cbo����.ItemData(Cbo����.ListIndex)
    End If
End Sub

Private Sub Cbo����_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cbo����_Validate(Cancel As Boolean)
    If Cbo����.ListCount > 0 Then
        If Cbo����.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub ChkҩƷ_Click()
    TxtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    CmdҩƷ.Enabled = TxtҩƷ.Enabled
    If TxtҩƷ.Enabled Then TxtҩƷ.SetFocus
End Sub

Private Sub CmdCancel_Click()
    strReturn = ""
    Unload Me
End Sub

Private Sub InitIDKindNew()
    Call IDKNType.zlInit(Me, glngSys, 1341, gcnOracle, gstrDbUser, mobjSquareCard, "", txt���￨, , True)
End Sub

Private Sub cmdOk_Click()
    If CheckData = False Then Exit Sub
    Call GetSQL
    
    FrmObj.intģʽ = IIf(BlnState, -1, 1)
    Unload Me
End Sub

Private Sub cmdҩƷ_Click()
    Dim RecReturn As New ADODB.Recordset
    
'    With FrmҩƷѡ����
'        Set RecReturn = .ShowME(Me, 1, lngҩ��ID, , , False)
'    End With
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ������ҩ", lngҩ��ID, lngҩ��ID)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 1, , , , lngҩ��ID, , , , False, , , , False)
    
    With RecReturn
        If .EOF Then Exit Sub
        TxtҩƷ.Tag = !ҩƷID
        TxtҩƷ = "[" & !ҩƷ���� & "]" & IIf(IsNull(!ͨ����), "", !ͨ����)
    End With
End Sub

Private Sub Form_Load()
    Dim intDays As Integer
    Dim dateCurDate As Date
    
    BlnStartUp = False
    BlnState = (IntOper = 6)
    strReturn = ""
    
    intDays = Val(zldatabase.GetPara("��ѯ����", glngSys, 1341, 1))
    intDays = intDays - 1
    
    dateCurDate = Sys.Currentdate()
    Me.Dtp��ʼDate.Value = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
    Me.Dtp����Date.Value = Format(dateCurDate, "yyyy-MM-dd 23:59:59")
    
    Select Case IntOper
    Case 1
        Me.Caption = "����δ��ҩ��������"
    Case 2
        Me.Caption = "��������ҩ��������"
    Case 3
        Me.Caption = "����δ��ҩ��������"
    Case 4
        Me.Caption = "���ҳ�ʱδ��ҩ��������"
    Case 5
        Me.Caption = "�����ѷ�ҩ��������"
    End Select
    
    If DependOnCheck = False Then Exit Sub
    If glngSys \ 100 <> 1 Then
        Lbl����.Visible = False
        Cbo����.Visible = False
    End If
    
    If IntOper <> 5 Then
        Me.Dtp��ʼDate.Enabled = zlStr.IsHavePrivs(mstrPrivs, "�޸Ĺ�������")
        Me.Dtp����Date.Enabled = Me.Dtp��ʼDate.Enabled
    End If
    
    If Not IsInString(mstrPrivs, "�����ѯ����ʱ�䷶Χ����", ";") Then
        Dtp��ʼDate.Enabled = False
        Dtp����Date.Enabled = False
    End If
    
    BlnStartUp = True
    
    Call zlfuncCard_Ini(mobjSquareCard, Me, 1341)
    Call InitIDKindNew
    
    On Error Resume Next
    If mbln���￨ Then txt���￨.SetFocus
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
End Sub

Private Sub IDKNType_ItemClick(index As Integer, objCard As zlIDKind.Card)
    If objCard.�������Ĺ��� <> "" Then
        txt���￨.PasswordChar = "*"
    Else
        txt���￨.PasswordChar = ""
    End If
End Sub

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txt���￨.Text = objPatiInfor.����
End Sub

Private Sub Txt����NO_GotFocus()
    GetFocus Txt����NO
End Sub

Private Sub Txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intYear As Integer, strYear As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt����NO) = "" Then Exit Sub
    '--���������λ,�򰴹������--
    Me.Txt����NO = GetFullNO(UCase(LTrim(Me.Txt����NO)), 13)
End Sub

Private Sub txt���￨_Change()
    If txt���￨.Text <> "" And Len(txt���￨.Text) = 18 And Not mobjSquareCard Is Nothing And IDKNType.GetCurCard.���� = "�������֤" Then
        If mobjSquareCard.zlGetPatiID("���֤", UCase(txt���￨.Text), False, mlng����ID) = False Then mlng����ID = 0
    End If
End Sub

Private Sub txt���￨_KeyPress(KeyAscii As Integer)
    'ȥ���ſ��������������ַ�
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub


Private Sub Txt��ʼNO_GotFocus()
    GetFocus Txt��ʼNO
End Sub

Private Sub Txt��ʼNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt��ʼNO) = "" Then Exit Sub
    '--���������λ,�򰴹������--
    Me.Txt��ʼNO = GetFullNO(UCase(LTrim(Me.Txt��ʼNO)), 13)
End Sub



Private Sub Txt����_GotFocus()
    GetFocus Txt����
End Sub

Private Sub TxtҩƷ_GotFocus()
    GetFocus TxtҩƷ
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.Height = IIf((BlnState And zlStr.IsHavePrivs(mstrPrivs, "���˸�������")), DblAdvanceHeight, DblNormalHeight)
    fra��������.Enabled = (BlnState And zlStr.IsHavePrivs(mstrPrivs, "���˸�������"))
    
    With fra��������
        .Top = IIf((BlnState And zlStr.IsHavePrivs(mstrPrivs, "���˸�������")), Fra��������.Top + Fra��������.Height + 80, CmdOK.Top + CmdOK.Height + 180)
    End With
    With CmdOK
        .Top = IIf((BlnState And zlStr.IsHavePrivs(mstrPrivs, "���˸�������")), fra��������.Top + fra��������.Height + 80, Fra��������.Top + Fra��������.Height + 80)
    End With
    With CmdCancel
        .Top = CmdOK.Top
    End With
    
    If fra��������.Enabled = True Then
        Me.Cbo������.Enabled = zlStr.IsHavePrivs(mstrPrivs, "ҽ����ѯ")
    End If
End Sub

Private Function DependOnCheck() As Boolean
    Dim RecTmp As ADODB.Recordset

    '������������Ƿ�����
    DependOnCheck = False
    
    On Error GoTo errHandle
    Cbo����.Clear
    
    If glngSys \ 100 = 1 Then
        '���ݵ�ǰ���ŵķ��������ȡ����
        gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
                 " Where ID in (Select ����ID From ��������˵�� Where �������� In ('�ٴ�','����') And ������� IN(1,2,3))" & _
                 " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                 " Order By ����||'-'||���� "
        Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, "DependOnCheck")

        If RecTmp.EOF Then
            MsgBox "���ʼ�����ű����Ź�����", vbInformation, gstrSysName
            Exit Function
        End If
        Me.Cbo����.AddItem "����"
        Do While Not RecTmp.EOF
            Cbo����.AddItem RecTmp!����
            Cbo����.ItemData(Cbo����.NewIndex) = RecTmp!Id
            RecTmp.MoveNext
        Loop
        Cbo����.ListIndex = 0
    End If
        
    If IntOper = 6 Then
        '���������
        Cbo������.Clear
        Cbo������.AddItem "����"

        gstrSQL = " Select distinct ���� ������ From ��Ա��" & _
                 " Where ID IN (" & _
                 " Select ��ԱID From ��Ա����˵��" & _
                 " Where ��Ա����='ҽ��')" & _
                 " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
         
        Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, "DependOnCheck")
        
        Do While Not RecTmp.EOF
            Cbo������.AddItem Trim(RecTmp!������)
            RecTmp.MoveNext
        Loop
        Cbo������.ListIndex = 0
        
        '��ӷ�ҩ��
        Cbo��ҩ��.Clear
        Cbo��ҩ��.AddItem "����"

        gstrSQL = " Select distinct ���� ����� From ��Ա��" & _
                 " Where ID IN (" & _
                 " Select ��ԱID From ��Ա����˵��" & _
                 " Where ��Ա����='ҩ����ҩ��')" & _
                 " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, "DependOnCheck")

        Do While Not RecTmp.EOF
            Cbo��ҩ��.AddItem Trim(RecTmp!�����)
            RecTmp.MoveNext
        Loop
        Cbo��ҩ��.ListIndex = 0
    End If

    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
    '���������ȷ��
    CheckData = False
    
    Txt��ʼNO = UCase(Trim(Txt��ʼNO))
    Txt����NO = UCase(Trim(Txt����NO))
    Txt���� = UCase(Trim(Txt����))
    
    If BlnState Then
'        Cbo������ = Trim(Cbo������)
'        Cbo��ҩ�� = Trim(Cbo��ҩ��)
        If ChkҩƷ.Value = 1 Then
            If TxtҩƷ.Tag = 0 Then
                MsgBox "������ҩƷ��Ϣ��", vbInformation, gstrSysName
                TxtҩƷ.SetFocus
                Exit Function
            End If
        End If
    End If
    
    CheckData = True
End Function

Private Function GetSQL()
    '�����û��������SQL
    strReturn = ""
    
    If BlnState = False Then
        strReturn = " And A.�������� Between [1] And [2] "
    Else
        strReturn = " And A.������� Between [1] And [2] "
    End If
    
    If Txt��ʼNO <> "" Or Txt����NO <> "" Then
        If Txt��ʼNO <> "" And Txt����NO <> "" Then
            strReturn = strReturn & " And A.NO Between [3] And [4] "
        Else
            If Txt��ʼNO <> "" Then
                strReturn = strReturn & " And A.NO = [3] "
            Else
                strReturn = strReturn & " And A.NO = [4] "
            End If
        End If
    End If
    
    If BlnState = False Then
        If Txt���� <> "" Then strReturn = strReturn & " And Upper(A.����) Like [5] "
        If txtסԺ�� <> "" Then strReturn = strReturn & " And Upper(DECODE(A.����,8,A.�����,A.סԺ��)) Like [7] "
        If Cbo����.ListIndex <> 0 And glngSys \ 100 = 1 Then strReturn = strReturn & " And C.�Է�����ID+0=[8] "
    Else
        If Txt���� <> "" Then strReturn = strReturn & " And Upper(H.����) Like [5] "
        If txtסԺ�� <> "" Then strReturn = strReturn & " And Upper(H.��ʶ��) Like [7] "
        If Cbo����.ListIndex <> 0 And glngSys \ 100 = 1 Then strReturn = strReturn & " And A.�Է�����ID+0=[8] "
    End If
    If Trim(txt���￨.Text) <> "" Then
        mbln���￨ = True
        If BlnState = False Then
            strReturn = strReturn & " And Upper(A.���￨��) = [6] "
        Else
            strReturn = strReturn & " And Upper(B.���￨��) = [6] "
        End If
    End If
    
    SQLCondition.date��ʼ���� = CDate(Format(Me.Dtp��ʼDate, "yyyy-MM-dd hh:mm:ss"))
    SQLCondition.date�������� = CDate(Format(Me.Dtp����Date, "yyyy-MM-dd hh:mm:ss"))
    SQLCondition.str��ʼNO = Txt��ʼNO
    SQLCondition.str����NO = Txt����NO
    SQLCondition.str���� = IIf(Txt���� = "", "", Txt���� & "%")
    SQLCondition.str���￨ = txt���￨.Text & IIf(txt���￨.Text = "", "", "|" & IDKNType.GetCurCard.���� & "," & IDKNType.GetCurCard.�ӿ���� & IIf(mlng����ID <> 0, "|" & mlng����ID, ""))
    SQLCondition.str��ʶ�� = IIf(txtסԺ�� = "", "", txtסԺ�� & "%")
    SQLCondition.lng����ID = Cbo����.ItemData(Cbo����.ListIndex)
    SQLCondition.strҽ���� = UCase(Trim(txtҽ����.Text))
    
    '��ҩ����
    If chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
        mint��Ժ��ҩ = 0
    ElseIf chkSend(0).Value = 1 Then
        mint��Ժ��ҩ = 1
    ElseIf chkSend(1).Value = 1 Then
        mint��Ժ��ҩ = 2
    End If
    
    If BlnState = False Then Exit Function
    
    If Cbo������.ListIndex <> 0 Then strReturn = strReturn & " And Trim(A.������) Like [9] "
    If Cbo��ҩ��.ListIndex <> 0 Then strReturn = strReturn & " And Trim(A.�����) Like [10] "
    If Val(TxtҩƷ.Tag) <> 0 Then strReturn = strReturn & " And A.ҩƷID+0=[11] "
    
    SQLCondition.str������ = ""
    SQLCondition.str����� = ""
    SQLCondition.lngҩƷid = 0
    
    If Cbo������.Text <> "" And Cbo������.Text <> "����" Then SQLCondition.str������ = Cbo������.Text
    If Cbo��ҩ��.Text <> "" And Cbo��ҩ��.Text <> "����" Then SQLCondition.str����� = Cbo��ҩ��.Text
    SQLCondition.lngҩƷid = Val(TxtҩƷ.Tag)
End Function

Public Function ShowMe(ByVal FrmMain As Form, ByVal In_Lngҩ��ID As Long, _
    ByVal In_Int����ģʽ As Integer, ByVal In_Ȩ�� As String, bln���￨ As Boolean, _
    ByRef date��ʼ���� As Date, _
    ByRef date�������� As Date, _
    ByRef str��ʼNO As String, _
    ByRef str����NO As String, _
    ByRef str���� As String, _
    ByRef str���￨ As String, _
    ByRef str��ʶ�� As String, _
    ByRef lng����ID As Long, _
    ByRef str������ As String, _
    ByRef str����� As String, _
    ByRef lngҩƷid As Long, _
    ByRef strҽ���� As String, _
    ByRef int��Ժ��ҩ As Integer) As String
    
    lngҩ��ID = In_Lngҩ��ID
    IntOper = In_Int����ģʽ
    mstrPrivs = In_Ȩ��
    mbln���￨ = bln���￨
    
    Set FrmObj = FrmMain
    With Me
        .Show 1, FrmMain
    End With
    
    bln���￨ = mbln���￨
    
    date��ʼ���� = SQLCondition.date��ʼ����
    date�������� = SQLCondition.date��������
    str��ʼNO = SQLCondition.str��ʼNO
    str����NO = SQLCondition.str����NO
    str���� = SQLCondition.str����
    str���￨ = SQLCondition.str���￨
    str��ʶ�� = SQLCondition.str��ʶ��
    lng����ID = SQLCondition.lng����ID
    str������ = SQLCondition.str������
    str����� = SQLCondition.str�����
    lngҩƷid = SQLCondition.lngҩƷid
    strҽ���� = SQLCondition.strҽ����
    int��Ժ��ҩ = mint��Ժ��ҩ
    
    ShowMe = strReturn
End Function

Private Sub TxtҩƷ_Validate(Cancel As Boolean)
    TxtҩƷ = Trim(TxtҩƷ)
    If TxtҩƷ = "" Then
        TxtҩƷ.Tag = 0
        Exit Sub
    End If
    
    Dim RecReturn As New ADODB.Recordset
    Dim sngLeft As Single, sngTop As Single
    
    If InStr(1, TxtҩƷ, "[") <> 0 And InStr(1, TxtҩƷ, "]") <> 0 Then TxtҩƷ.Text = Mid(TxtҩƷ.Text, 2, InStr(1, TxtҩƷ, "]") - 2)
    sngLeft = Me.Left + TxtҩƷ.Left + fra��������.Left + 50
    sngTop = Me.Top + (Me.Height - Me.ScaleHeight) + TxtҩƷ.Top + fra��������.Top + TxtҩƷ.Height - 100
    If DblFrmHeight + sngTop > Screen.Height Then sngTop = sngTop - DblFrmHeight - TxtҩƷ.Height + 50
'    With FrmҩƷ��ѡѡ����
'        Set RecReturn = .ShowME(Me, 1, lngҩ��ID, , , TxtҩƷ.Text, sngLeft, sngTop, False)
'        If RecReturn.EOF Then Cancel = True: Exit Sub
'    End With
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ������ҩ", lngҩ��ID, lngҩ��ID)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 1, 1, UCase(TxtҩƷ.Text), sngLeft, sngTop, lngҩ��ID, , , , False, , , , False)
    
    If RecReturn.EOF Then Cancel = True: Exit Sub
    TxtҩƷ.Tag = RecReturn!ҩƷID
    TxtҩƷ = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!ͨ����), "", RecReturn!ͨ����)
End Sub
