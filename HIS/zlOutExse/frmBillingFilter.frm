VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmBillingFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   5295
      TabIndex        =   14
      Top             =   1485
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   105
      TabIndex        =   15
      Top             =   0
      Width           =   5010
      Begin VB.TextBox txtPatientNo 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   18
         TabIndex        =   10
         Top             =   2325
         Width           =   3825
      End
      Begin VB.TextBox txtIdentify 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1575
         MaxLength       =   64
         TabIndex        =   11
         Top             =   2730
         Width           =   3225
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���ʵ���"
         Height          =   210
         Left            =   3255
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1515
         Width           =   1545
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   975
         TabIndex        =   8
         Text            =   "cbo����"
         Top             =   1920
         Width           =   1545
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���ʵ���"
         Height          =   210
         Left            =   3255
         TabIndex        =   3
         Top             =   705
         Width           =   1020
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3195
         TabIndex        =   9
         Text            =   "cbo����Ա"
         Top             =   1920
         Width           =   1590
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1098
         Width           =   1545
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3195
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1098
         Width           =   1590
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3195
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1515
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   684
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
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   975
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
         Left            =   975
         TabIndex        =   24
         Top             =   2730
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Appearance      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lblPatientNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   135
         TabIndex        =   26
         Top             =   2385
         Width           =   765
      End
      Begin VB.Label lblIdentity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʶ��"
         Height          =   180
         Left            =   180
         TabIndex        =   25
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   180
         Left            =   360
         TabIndex        =   23
         Top             =   1575
         Width           =   540
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   744
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   2805
         TabIndex        =   20
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   360
         TabIndex        =   19
         Top             =   1158
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2805
         TabIndex        =   18
         Top             =   1575
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   1986
         Width           =   720
      End
      Begin VB.Label lbl����Ա 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ա"
         Height          =   180
         Left            =   2625
         TabIndex        =   16
         Top             =   1980
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5295
      TabIndex        =   13
      Top             =   675
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5295
      TabIndex        =   12
      Top             =   255
      Width           =   1100
   End
End
Attribute VB_Name = "frmBillingFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrFilter As String
Public mblnDateMoved As Boolean 'Out
Public mstrPrivs As String
Private mrsPerson As ADODB.Recordset
Private Const mlngModule = 1122
Private mlngPreID As Long
Public mlngPrePatient As Long '�����:38539
Private mrsInfo As ADODB.Recordset '�����:38539
Private mblnOlnyBJYB As Boolean '�����:38539
Private mblnKeyReturn As Boolean '�����:38539
Private mblnNotClick As Boolean '�����:38539
Private mblnUnChange  As Boolean '�����:38539
Private mrsDept As ADODB.Recordset

Private Sub cbo����Ա_Click()
    If cbo����Ա.ListIndex >= 0 Then mlngPreID = cbo����Ա.ItemData(cbo����Ա.ListIndex)
End Sub

Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
   Dim lngIdx As Long, lngҽ��ID As Long
    Dim strAllCaption As String
    
    '���˺� ����:21899
    If KeyAscii <> 13 Then Exit Sub
    If cbo����Ա.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If InStr(1, mstrPrivs, ";���в���Ա;") = 0 Then
        cbo����Ա.ListIndex = 0: Exit Sub
    End If
    strAllCaption = "���в���Ա"
    
    If mrsPerson Is Nothing Then Exit Sub
    If zlPersonSelect(Me, mlngModule, cbo����Ա, mrsPerson, cbo����Ա.Text, True, strAllCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
    

'    Dim lngIdx As Long
'    If KeyAscii >= 32 Then
'        lngIdx = zlControl.CboMatchIndex(cbo����Ա.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo����Ա.ListCount > 0 Then lngIdx = 0
'        cbo����Ա.ListIndex = lngIdx
'    End If
End Sub

Private Sub cbo����Ա_Validate(Cancel As Boolean)
    
    If cbo����Ա.ListIndex < 0 Then zlControl.CboLocate cbo����Ա, mlngPreID, True
    If cbo����Ա.ListIndex < 0 And cbo����Ա.Text <> "" Then cbo����Ա.ListIndex = 0
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
    If zlSelectDept(Me, 1120, cbo����, mrsDept, cbo����.Text, True, "���п���") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub chk����_Click()
    If chk����.Enabled And chk����.Enabled Then
        If chk����.Value = 0 And chk����.Value = 0 Then
            chk����.Value = 1
        End If
    End If
End Sub

Public Sub chk����_Click()
    If chk����.Enabled And chk����.Enabled Then
        If chk����.Value = 0 And chk����.Value = 0 Then
            chk����.Value = 1
        End If
    End If
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
    If Me.ActiveControl Is cbo����Ա Then Exit Sub
    If Me.ActiveControl Is cbo���� Then Exit Sub
    
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Me.ActiveControl Is cbo����Ա Then Exit Sub
    If Me.ActiveControl Is cbo���� Then Exit Sub
    
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngOldID As Long
    On Error GoTo errH
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txt����ID.Text = ""
    txt����.Text = ""
    chk����.Value = 1
    chk����.Value = 0
    
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    '����Ա
    cbo����Ա.Clear
    If InStr(mstrPrivs, "���в���Ա") > 0 Then  '21899
            cbo����Ա.AddItem "���в���Ա"
            cbo����Ա.ListIndex = 0
            Set mrsPerson = GetPersonnel("", True)
            For i = 1 To mrsPerson.RecordCount
                cbo����Ա.AddItem mrsPerson!���� & "-" & mrsPerson!����
                cbo����Ա.ItemData(cbo����Ա.NewIndex) = mrsPerson!ID
                mrsPerson.MoveNext
            Next
    Else
        cbo����Ա.AddItem UserInfo.���� & "-" & UserInfo.����
        cbo����Ա.ItemData(cbo����Ա.NewIndex) = UserInfo.ID
    End If
    If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount > 0 Then cbo����Ա.ListIndex = 0
    
    '��������
    cbo����.Clear
    cbo����.AddItem "���п���"
    cbo����.ListIndex = 0
    Set mrsDept = GetDepartments("'�ٴ�','����'", "1,3")
    For i = 1 To mrsDept.RecordCount
        If lngOldID <> mrsDept!ID Then
            cbo����.AddItem mrsDept!���� & "-" & mrsDept!����
            cbo����.ItemData(cbo����.NewIndex) = mrsDept!ID
            lngOldID = mrsDept!ID
        End If
        mrsDept.MoveNext
    Next
    
    '�����:38539
    InitIDKind
    
    Call chk����_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsDept Is Nothing Then Set mrsDept = Nothing
End Sub

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
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 14)
End Sub
Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 14)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Public Sub MakeFilter()
    mstrFilter = " And �Ǽ�ʱ�� Between [1] And [2]"
    
    If chk����.Enabled = True Then
        mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    Else
        '���۵�ɸѡʱ,���ôӺ����ݱ�ȡ
        mblnDateMoved = False
    End If
        
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3]"
    End If
    
    If txt����.Text <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt����.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(����) Like [5]"
        Else
            mstrFilter = mstrFilter & " And ���� Like [5]"
        End If
    End If
    
    If IsNumeric(txt����ID.Text) Then
        mstrFilter = mstrFilter & " And ����ID=[6]"
    End If
    
    If cbo����.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And ��������ID+0=[7]"
    End If
    
    '�����:38539
    If txtPatientNo.Text <> "" Then mstrFilter = mstrFilter & " And ��ʶ��=[8]"
    '�����:38539
    If txtIdentify.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
            If Val(Nvl(mrsInfo!ID)) = mlngPrePatient Then
                mstrFilter = mstrFilter & " And ����ID=[9]"
            End If
    End If
    
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����ID_GotFocus()
    zlControl.TxtSelAll txt����ID
End Sub

'------------------------------------------------------------

Private Sub GetPatient(ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '�����:38539
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
    
    strSQL = ""
    mlngPrePatient = 0
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) And InStr("-+*", Left(strInput, 1)) = 0 Then       '103563
       
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
                    If txtIdentify.Text = mrsInfo!���� Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtIdentify.Text = ""
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
'
'     vRect = zlcontrol.GetControlRect(txtIdentify.hWnd)
'     Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "��", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtIdentify.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%")
     If Not mrsInfo Is Nothing Then
        If mrsInfo.RecordCount = 0 Then
            Set mrsInfo = Nothing
            txtIdentify.Text = ""
            Exit Sub
        End If
        If mrsInfo!ID = 0 Then  'û���ҵ�������Ϣ
            Set mrsInfo = Nothing
            txtIdentify.Text = ""
            Exit Sub
        Else '��ȡ��������Ϣ
        
          txtIdentify.Text = Nvl(mrsInfo!����)
          Me.txtIdentify.Tag = Nvl(mrsInfo!ID)
          mlngPrePatient = Val(Nvl(mrsInfo!ID))
         
        End If
    Else 'ȡ��ѡ��
        txtIdentify.Text = ""
        Set mrsInfo = Nothing: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub txtIdentify_Change()
'�����:38539
    txtIdentify.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtIdentify Then
        IDKind.SetAutoReadCard txtIdentify.Text = ""
    End If
   
End Sub


Private Sub txtIdentify_GotFocus()
'�����:38539
    Call zlControl.TxtSelAll(txtIdentify)
    Call zlCommFun.OpenIme(True)
    If txtIdentify.Text = "" And ActiveControl Is txtIdentify Then IDKind.SetAutoReadCard True
End Sub


Private Sub txtIdentify_LostFocus()
'�����:38539
    IDKind.SetAutoReadCard False
End Sub

Private Sub txtIdentify_Validate(Cancel As Boolean)
'�����:38539
    If mblnKeyReturn = False Then
        Call txtIdentify_KeyPress(13)
    Else
        mblnKeyReturn = False
    End If
End Sub

Private Sub txtIdentify_KeyPress(KeyAscii As Integer)
'�����:38539
  Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
        On Error GoTo errH
        If txtIdentify.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IsCardType(IDKind, "����") Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtIdentify.Text, 1)) > 0 And IsNumeric(Mid(txtIdentify.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtIdentify, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    End If
    If blnCard And Len(txtIdentify.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtIdentify.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtIdentify.Text = txtIdentify.Text & Chr(KeyAscii)
            txtIdentify.SelStart = Len(txtIdentify.Text)
        ElseIf IsNumeric(txtIdentify.Tag) Then
            KeyAscii = 0
            'If txtIdentify.Tag <> "" Then
            'ˢ�²�����Ϣ:"-����ID"
            If Val(txtIdentify.Tag) <> 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            Call GetPatient(txtIdentify.Tag, False)
            Exit Sub
        End If
        KeyAscii = 0
        If IsCardType(IDKind, "IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtIdentify.Text, 1)) = 0)
        Call GetPatient(txtIdentify.Text, blnCard)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog '
End Sub


'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
'�����:38539
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtIdentify)
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
'�����:38539
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
'�����:38539
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
'�����:38539
    '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    Set gobjSquare.objCurCard = objCard
    
    txtIdentify.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtIdentify.IMEMode = 0
    
     '��������ʾ,Ҳ�����г�������,���ﲻ�漰���밲ȫ,ֻ�����������Ϣ��ȡ
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtIdentify.Text <> "" And Not mblnNotClick Then txtIdentify.Text = ""
    If txtIdentify.Enabled And txtIdentify.Visible Then txtIdentify.SetFocus
    If mlngPrePatient Then txtIdentify.PasswordChar = ""
    zlControl.TxtSelAll txtIdentify
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
'�����:38539
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtIdentify.Locked Then Exit Sub
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
'        'ϵͳIC��
'        If Not mobjICCard Is Nothing Then
'           txtIdentify.Text = mobjICCard.Read_Card()
'           If txtIdentify.Text <> "" Then
'                mblnUnChange = True
'                Call txtIdentify_Validate(False)
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
    txtIdentify.Text = strOutCardNO
    
    If txtIdentify.Text <> "" Then
        mblnUnChange = True
        Call txtIdentify_Validate(False)
        mblnUnChange = False
    End If
    
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
'�����:38539
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
     '����:60010
    If txtIdentify.Locked Then Exit Sub   'Or Not Me.ActiveControl Is txtIdentify
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    If IsCardType(IDKind, "���֤") Then
        txtIdentify.Text = objPatiInfor.���֤��
    Else
        txtIdentify.Text = objPatiInfor.����
    End If
    Call txtIdentify_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub


Private Sub txtPatientNo_KeyPress(KeyAscii As Integer)
'�����:38539
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

