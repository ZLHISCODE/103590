VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPriceFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra�շ���� 
      Caption         =   "�շ����"
      Height          =   2295
      Left            =   30
      TabIndex        =   25
      Top             =   3180
      Width           =   5775
      Begin MSComctlLib.ListView lvwType 
         Height          =   2025
         Left            =   60
         TabIndex        =   13
         Top             =   210
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   3572
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   5970
      TabIndex        =   16
      Top             =   1455
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3120
      Left            =   30
      TabIndex        =   17
      Top             =   0
      Width           =   5790
      Begin VB.OptionButton opt���� 
         Caption         =   "���ﲡ�˺�סԺ����"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Top             =   2790
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "סԺ����"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2220
         TabIndex        =   11
         Top             =   2790
         Width           =   1020
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "���ﲡ��"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   2790
         Width           =   1020
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3600
         MaxLength       =   64
         TabIndex        =   6
         Top             =   1500
         Width           =   2070
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1440
         MaxLength       =   64
         TabIndex        =   9
         Top             =   2340
         Width           =   4215
      End
      Begin VB.CheckBox chk�շ� 
         Caption         =   "�����շѵ���"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   713
         Width           =   1440
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1932
         Width           =   2070
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   855
         TabIndex        =   5
         Text            =   "cbo����"
         Top             =   1500
         Width           =   2070
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   2070
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1098
         Width           =   2070
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1098
         Width           =   2070
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   855
         TabIndex        =   1
         Top             =   690
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   855
         TabIndex        =   0
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   840
         TabIndex        =   26
         Top             =   2340
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   "��|���￨|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|0|0|0|0|0|;IC|IC����|1|0|0|0|0|;��|�����|0|0|0|0|0|"
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
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lblFil 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������Դ"
         Height          =   180
         Left            =   60
         TabIndex        =   29
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3180
         TabIndex        =   28
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʶ��"
         Height          =   180
         Left            =   60
         TabIndex        =   27
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   3180
         TabIndex        =   24
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   60
         TabIndex        =   23
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   60
         TabIndex        =   22
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3180
         TabIndex        =   21
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   60
         TabIndex        =   19
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lbl����Ա 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   1980
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5970
      TabIndex        =   15
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5970
      TabIndex        =   14
      Top             =   225
      Width           =   1100
   End
End
Attribute VB_Name = "frmPriceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrFilter As String
Public mblnDateMoved As Boolean
Public mstrPrivs As String
Public mstr�շ���� As String


Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnNotClick As Boolean
Private mblnUnChange  As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mrsDept As ADODB.Recordset

Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo����Ա.hWnd, KeyAscii)
        If lngIdx = -1 And cbo����Ա.ListCount > 0 Then lngIdx = 0
        cbo����Ա.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo�ѱ�.hWnd, KeyAscii)
        If lngIdx = -1 And cbo�ѱ�.ListCount > 0 Then lngIdx = 0
        cbo�ѱ�.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo����_Click()
    Dim lng��������ID As Long
    
    On Error GoTo errHandler
    If cbo����.ListIndex <> -1 Then lng��������ID = cbo����.ItemData(cbo����.ListIndex)
    If Val(cbo����.Tag) = lng��������ID Then Exit Sub
    cbo����.Tag = lng��������ID
        
    '��λҽ��
    If gbyt����ҽ�� = 1 Then
        If cbo����.ListIndex <> -1 Then
            Call FillPerson(lng��������ID)
        Else
            cbo����Ա.Clear
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillPerson(Optional ByVal lng����ID As Long)
'���ܣ�����ָ���Ŀ�������ID��ȡ����дҽ���б�,����ȱʡҽ��
    Dim rsTmp As ADODB.Recordset
    Dim bln������Ա���� As Boolean
    
    cbo����Ա.Clear
    cbo����Ա.AddItem "���л�����"
    bln������Ա���� = zlStr.IsHavePrivs(mstrPrivs, "���п���") = False And gblnUserIsClinic '113577
    Set rsTmp = GetPersonnel("ҩ����ҩ��,ҽ��,��ʿ", True, bln������Ա����, lng����ID)
    Do While Not rsTmp.EOF
        cbo����Ա.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ID = UserInfo.ID Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
        rsTmp.MoveNext
    Loop
    If cbo����Ա.ListIndex < 0 And cbo����Ա.ListCount > 0 Then cbo����Ա.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
'    If KeyAscii >= 32 Then
'        lngIdx = zlControl.CboMatchIndex(cbo����.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo����.ListCount > 0 Then lngIdx = 0
'        cbo����.ListIndex = lngIdx
'    End If
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo����.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Set mrsDept = GetDepartments("'�ٴ�','����'", gint������Դ & ",3")
    If zlSelectDept(Me, 1120, cbo����, mrsDept, cbo����.Text, True, _
        IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), "���п���", "")) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    If cbo����.ListIndex >= 0 Then Exit Sub
    If cbo����.ListIndex < 0 And cbo����.ListCount <> 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "�������ݺŲ���С�ڿ�ʼ���ݺţ�", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    Call zlGet�շ����
    If mstr�շ���� = "" Then
        MsgBox "����ѡ��һ�������в���,��ѡ�����!", vbInformation + vbOKOnly, gstrSysName
        If lvwType.Enabled Then lvwType.SetFocus
        Exit Sub
    End If
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl Is cbo���� Then Exit Sub
    
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Me.ActiveControl Is cbo���� Then Exit Sub
    
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Integer, lngOldID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim bln������Ա���� As Boolean
    On Error GoTo errH
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtPatient.Text = ""
    
    chk�շ�.Value = 0
    Call InitIDKind
    '���ó�ʼֵ
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    '����Ա
    Call FillPerson
    cbo.SetListWidth cbo����Ա.hWnd, cbo����Ա.Width * 3 / 2
    
    '��ѡ�ѱ�
    cbo�ѱ�.Clear
    cbo�ѱ�.AddItem "���зѱ�"
    cbo�ѱ�.ListIndex = 0
    
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Where Nvl(�������,3) IN(1,3) Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    '��������
    bln������Ա���� = zlStr.IsHavePrivs(mstrPrivs, "���п���") = False And gblnUserIsClinic '113577
    cbo����.Clear: cbo����.Tag = ""
    If bln������Ա���� = False Then cbo����.AddItem "���п���"
    Set mrsDept = GetDepartments("'�ٴ�','����','����','���','����','����'", gint������Դ & ",3", bln������Ա����)
    For i = 1 To mrsDept.RecordCount
        If lngOldID <> mrsDept!ID Then
            cbo����.AddItem mrsDept!���� & "-" & mrsDept!����
            cbo����.ItemData(cbo����.NewIndex) = mrsDept!ID
            lngOldID = mrsDept!ID
        End If
        mrsDept.MoveNext
    Next
    If cbo����.ListIndex < 0 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
    cbo.SetListWidth cbo����.hWnd, cbo����.Width * 3 / 2
    
    Dim str�շ���� As String
    str�շ���� = zlDatabase.GetPara("�ϴι����շ����", glngSys, 1120, "", Array(lvwType, fra�շ����), InStr(1, mstrPrivs, "��������") > 0)
    
    strSQL = "Select ����,���� as ��� from �շ���Ŀ���  Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Dim objList As ListItem
    With lvwType
        .ListItems.Clear
        Do While Not rsTmp.EOF
            Set objList = .ListItems.Add(, "K" & Nvl(rsTmp!����), Nvl(rsTmp!���))
            If str�շ���� = "" Then
                objList.Checked = True
            Else
                objList.Checked = InStr(1, "," & str�շ���� & ",", "," & rsTmp!���� & ",") > 0
            End If
            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objList As ListItem
    Dim str�շ���� As String
    str�շ���� = ""
    With lvwType
        For Each objList In .ListItems
                If objList.Checked Then
                    str�շ���� = str�շ���� & "," & Mid(objList.Key, 2)
                End If
        Next
    End With
    If str�շ���� <> "" Then str�շ���� = Mid(str�շ����, 2)
    zlDatabase.SetPara "�ϴι����շ����", str�շ����, glngSys, 1120, InStr(1, mstrPrivs, "��������") > 0
    If Not mrsDept Is Nothing Then Set mrsDept = Nothing
End Sub
Private Function zlGet�շ����() As String
    Dim objList As ListItem
    mstr�շ���� = ""
    With lvwType
        For Each objList In .ListItems
                If objList.Checked Then
                    mstr�շ���� = mstr�շ���� & "," & Mid(objList.Key, 2)
                End If
        Next
    End With
    If mstr�շ���� <> "" Then mstr�շ���� = Mid(mstr�շ����, 2)

End Function
Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '46516
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 13)
End Sub

Private Sub txtNOEnd_LostFocus()
 
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 13)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub
Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Private Sub MakeFilter()

    mstrFilter = " And �Ǽ�ʱ�� Between [1] And [2]"
    
    If chk�շ�.Value = 1 Then
        mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    Else
        mblnDateMoved = False
    End If
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3]"
    End If
    
    If cbo����Ա.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And ������||''=[5]"
    End If
    
    If txtPatient.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
        If Val(Nvl(mrsInfo!ID)) = mlngPrePatient Then
            mstrFilter = mstrFilter & " And ����ID=[11]"
        End If
    End If
    
    If txt����.Text <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt����.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(����) Like [6]"
        Else
            mstrFilter = mstrFilter & " And ���� Like [6]"
        End If
    End If
    
    
    If cbo�ѱ�.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And �ѱ�=[7]"
    End If
    
    If cbo����.ListIndex = 0 Then
        '��һ�����������п���
        If Val(cbo����.ItemData(cbo����.ListIndex)) > 0 Then
            mstrFilter = mstrFilter & " And ��������ID+0=[8]"
        End If
    ElseIf cbo����.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And ��������ID+0=[8]"
    End If
    If mstr�շ���� <> "" Then
        mstrFilter = mstrFilter & " And instr( [10], ','||�շ����||',')>0 "
    End If
End Sub

 


'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function
'��ȡĬ��IDKind����
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

 
'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.����
            Else
                If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
            End If
     End Select
End Function
                
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    Set gobjSquare.objCurCard = objCard
    '��7λ��,��ֻ��������,��Ȼȡ������
     
    txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    
    If IDKind.GetCardNoLen <> 0 Then
        txtPatient.MaxLength = IDKind.GetCardNoLen
    Else
        txtPatient.MaxLength = 64
    End If
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    If mlngPrePatient Then txtPatient.PasswordChar = ""
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
'        'ϵͳIC��
'        If Not mobjICCard Is Nothing Then
'           txtPatient.Text = mobjICCard.Read_Card()
'           If txtPatient.Text <> "" Then
'                mblnUnChange = True
'                Call txtPatient_Validate(False)
'                mblnUnChange = False
'           End If
'        End If
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
    If gobjSquare.objSquareCard.zlReadCard(Me, glngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        mblnUnChange = True
        Call txtPatient_Validate(False)
        mblnUnChange = False
    End If
    
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
     
    If IsCardType(IDKind, "���֤") Then
        txtPatient.Text = objPatiInfor.���֤��
    Else
        txtPatient.Text = objPatiInfor.����
    End If
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub
Private Sub GetPatient(ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    mlngPrePatient = 0
    strSQL = ""
    
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
        Else
            lng�����ID = -1
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        If lng����ID <= 0 Then lng����ID = 0
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.����ID=[2] " & str����Ժ
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And B.�����=[2]" & str����Ժ
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��(������Ժ)
        strSQL = strSQL & " And B.סԺ��=[2]" & str����Ժ
    Else
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!���� Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtPatient.Text = ""
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                       strSQL = strSQL & " And  B.���� Like [3]"
                       
                    End If
                Else
                    strSQL = strSQL & " And B.����ID=[2]"
                    strInput = "-" & Val(mrsInfo!����ID)
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '������ҽ������Ч:������:����:26982
                    strSQL = strSQL & " And B.ҽ���� like [3] " & str����Ժ
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.ҽ����=[1]" & str����Ժ
                End If
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                ' strSQL = strSQL & " And B.���֤��=[1] " & str����Ժ
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.�����=[1]" & str����Ժ
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.סԺ��=[1]" & str����Ժ
            Case Else
                '��������,��ȡ��صĲ���ID
                If Val(IDKind.GetCurCard.�ӿ����) >= 0 Then
                    lng�����ID = Val(IDKind.GetCurCard.�ӿ����)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                    If lng����ID = 0 Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID <= 0 Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select distinct  B.����id As ID, Decode(sign(nvl(ylkxx.����id,0)),0,'','��') as �����˻�, B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
    strSQL = strSQL & vbNewLine & "      A.���� ��������"
    strSQL = strSQL & vbNewLine & " From ������Ϣ B, ������� A,ҽ�ƿ���� YLK,����ҽ�ƿ���Ϣ YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.���� = A.���(+) and b.����id=ylkxx.����id(+) and ylkxx.״̬(+)=0 and  ylkxx.�����id=ylk.id(+)  and ylk.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
     
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
'     vRect = zlcontrol.GetControlRect(txtPatient.hWnd)
'     Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "��", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%")
    If Not mrsInfo Is Nothing Then ' And Not blnCancel
       If mrsInfo.RecordCount = 0 Then
            Set mrsInfo = Nothing
            txtPatient.Text = ""
            Exit Sub
        End If
        
        If mrsInfo!ID = 0 Then 'û���ҵ�������Ϣ
            Set mrsInfo = Nothing
            txtPatient.Text = ""
            Exit Sub
        Else '��ȡ��������Ϣ
        
          txtPatient.Text = Nvl(mrsInfo!����)
          Me.txtPatient.Tag = Nvl(mrsInfo!ID)
          mlngPrePatient = Val(Nvl(mrsInfo!ID))
         
        End If
    Else 'ȡ��ѡ��
        txtPatient.Text = ""
        Set mrsInfo = Nothing: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub txtPatient_Change()
    txtPatient.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtPatient Then
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
   
End Sub


Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then IDKind.SetAutoReadCard True
End Sub


Private Sub txtPatient_LostFocus()
    IDKind.SetAutoReadCard False
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnKeyReturn = False Then
        Call txtPatient_KeyPress(13)
    Else
        mblnKeyReturn = False
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
  Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IsCardType(IDKind, "����") Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) Then
            KeyAscii = 0
            'If txtPatient.Tag <> "" Then
            'ˢ�²�����Ϣ:"-����ID"
            If Val(txtPatient.Tag) <> 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            Call GetPatient(txtPatient.Tag, False)
            Exit Sub
        End If
        KeyAscii = 0
        If IsCardType(IDKind, "IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        Call GetPatient(txtPatient.Text, blnCard)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog '
End Sub

