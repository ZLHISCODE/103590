VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmDistFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   18
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5205
      TabIndex        =   19
      Top             =   2565
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   315
         Left            =   1020
         TabIndex        =   22
         Top             =   1890
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
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
      Begin VB.TextBox txtValue 
         Height          =   300
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "��λF3"
         Top             =   1890
         Width           =   4515
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         TabIndex        =   12
         Top             =   1094
         Width           =   2085
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   10
         Top             =   1094
         Width           =   2085
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   8
         TabIndex        =   8
         Top             =   682
         Width           =   2085
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   6
         Top             =   682
         Width           =   2085
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1485
         Width           =   2085
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1506
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3975
         TabIndex        =   4
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   140378115
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1005
         TabIndex        =   2
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   140378115
         CurrentDate     =   36588
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���˲���"
         Height          =   180
         Left            =   210
         TabIndex        =   21
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3420
         TabIndex        =   11
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label lbl����Ա 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һ�Ա"
         Height          =   180
         Left            =   3390
         TabIndex        =   15
         Top             =   1545
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   585
         TabIndex        =   13
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һ�ʱ��"
         Height          =   180
         Left            =   225
         TabIndex        =   1
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   405
         TabIndex        =   5
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3420
         TabIndex        =   7
         Top             =   735
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3420
         TabIndex        =   3
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   150
      TabIndex        =   20
      Top             =   2580
      Width           =   1100
   End
   Begin VB.Menu mnuIDKind 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu mnuIDKinds 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDistFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlngModul As Long
Public mstrFilter As String
Public mstrSectName As String   '����ָ����ǰĬ�ϵĿ���
Private mstrPrivs As String
Private mrsDept As ADODB.Recordset  '��¼�ٴ�����
Private mrs�Һ�Ա As ADODB.Recordset
Private mcllFiter As Variant       '������Ϣ
Private mblnOk As Boolean
Private mlngPrePatient As Long
Private mrsInfo As ADODB.Recordset
Private mblnKeyReturn As Boolean
Private mblnOlnyBJYB As Boolean
'-----------------------------------------------------

Public Function zlShowMe(ByVal frmMain As Form, ByVal lngModule As Long, _
    ByRef cllFilter As Variant, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������,��ȡ�����������
    '��Σ�frmMain-������
    '         lngModule-ģ���
    '���Σ�cllFilter-������ص�������Ϣ
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-02 15:25:35
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModule: Set mcllFiter = cllFilter: mblnOk = False
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
    If mblnOk Then Set cllFilter = mcllFiter
    zlShowMe = mblnOk
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ػ�������
    '���ƣ����˺�
    '���ڣ�2010-06-02 15:59:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim str�Һ�Ա As String, lng����ID As Long, i  As Long, strTmp As String
    
    If mrs�Һ�Ա Is Nothing Then
        Set mrs�Һ�Ա = GetPersonnel("����Һ�Ա", True)
    ElseIf mrs�Һ�Ա.State <> 1 Then
        Set mrs�Һ�Ա = GetPersonnel("����Һ�Ա", True)
    End If
    If Not mcllFiter Is Nothing Then
        str�Һ�Ա = Trim(mcllFiter("�Һ�Ա"))
        lng����ID = Val(mcllFiter("����"))
    End If
    '�Һ�Ա
    cbo����Ա.Clear
    cbo����Ա.AddItem "���йҺ�Ա"
    cbo����Ա.ListIndex = 0
    If mrs�Һ�Ա.RecordCount > 0 Then
        Call mrs�Һ�Ա.MoveFirst
        For i = 1 To mrs�Һ�Ա.RecordCount
            cbo����Ա.AddItem mrs�Һ�Ա!���� & "-" & mrs�Һ�Ա!����
            If str�Һ�Ա = Nvl(mrs�Һ�Ա!����) Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
            mrs�Һ�Ա.MoveNext
        Next
    End If
   '��ȡ�����ٴ����ң�����Ѿ���ȡ�Ͳ��ٶ�ȡ
    '143274:���ϴ�,2019/7/26���������Ա�����С����п��ҡ�Ȩ�ޣ�ֻ��ʾ����Ա��������
    strTmp = Get�������(glngSys, mlngModul, mstrPrivs)
    If strTmp = "" Then strTmp = UserInfo.����ID
    
    If mrsDept Is Nothing Then
        Set mrsDept = GetDepartments("'�ٴ�'", "1,3")
    ElseIf mrsDept.State <> 1 Then
        Set mrsDept = GetDepartments("'�ٴ�'", "1,3")
    End If
    
    cbo����.Clear
    cbo����.AddItem "���п���"
    cbo����.ListIndex = 0
    With mrsDept
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, "," & strTmp & ",", "," & !id & ",") > 0 Then
                cbo����.AddItem !���� & "-" & !����
                cbo����.ItemData(cbo����.NewIndex) = !id
                If lng����ID = Val(Nvl(!id)) Then cbo����.ListIndex = cbo����.NewIndex
            End If
            .MoveNext
        Loop
    End With
    LoadData = True
End Function
Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����Ա.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����Ա.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����Ա.ListIndex = lngIdx
    If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount <> 0 Then cbo����Ա.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount <> 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub cmdDef_Click()
    Dim Curdate As Date
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txtValue.Text = ""
    '������
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.����ʱ�� Between [1] And [2]"
    Set mcllFiter = Nothing
    Call InitCllData
    Call LoadData
End Sub

Private Sub cmdOK_Click()
    If Not IsNull(dtpEnd.Value) Then
        If dtpEnd.Value < dtpBegin.Value Then
            MsgBox "����ʱ�䲻��С�ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            dtpEnd.SetFocus: Exit Sub
        End If
    End If
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "�������ݺŲ���С�ڿ�ʼ���ݺţ�", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "����Ʊ�ݺŲ���С�ڿ�ʼƱ�ݺţ�", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    If MakeFilter = False Then Exit Sub
    mblnOk = True: Unload Me
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Not ActiveControl Is txtValue Then Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyF3 Then Call txtValue.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����:30346
    If InStr(1, "������������|��������<>?:;|'{}[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And Not ActiveControl Is txtValue Then KeyAscii = 0
End Sub

Public Sub Form_Load()
    Dim Curdate As Date, i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txtValue.Text = ""
    
    '������
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.����ʱ�� Between [1] And [2]"
    Call InitIDKind
    Call LoadData
    Call InitCllData
End Sub

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, rsTmp As ADODB.Recordset
    Dim lngCardID As Long, strSQL As String
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtValue)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0))
    '72936:������,2014-05-13,ȱʡ�������ͱ�ͣ�ú󱨴������
    If lngCardID <> 0 Then
        strSQL = "Select 1 From ҽ�ƿ���� Where ID=[1] And Nvl(�Ƿ�����,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardID)
        If Not rsTmp.EOF Then IDKind.DefaultCardType = lngCardID
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

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    mlngPrePatient = 0
    IDKind.SetAutoReadCard False
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
'        If mobjICCard Is Nothing Then
'            Set mobjICCard = CreateObject("zlICCard.clsICCard")
'            Set mobjICCard.gcnOracle = gcnOracle
'        End If
'        If mobjICCard Is Nothing Then Exit Sub
'        txtValue.Text = mobjICCard.Read_Card()
'        If txtValue.Text <> "" Then
'            Call FindPati(objCard, True, txtValue.Text)
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtValue.Text = strOutCardNO
    If txtValue.Text <> "" Then
        Call FindPati(objCard, True, txtValue.Text)
    End If
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    If txtValue.Text <> "" Then txtValue.Text = ""
    If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
    zlControl.TxtSelAll txtValue
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtValue.Locked Then Exit Sub
    txtValue.Text = objPatiInfor.����
    Call FindPati(objCard, True, txtValue.Text)
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
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
    '46512
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ

End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 12)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 12)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub


Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
   zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Private Function MakeFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĺ�������
    '����:���˺�
    '����:2011-10-21 15:23:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTmp As String, strSQLtmp As String
    Dim lng����ID As Long, lng�����ID As Long, strErrMsg As String, strPassWord As String
    Dim strKind As String
    Set mcllFiter = New Collection
    mstrFilter = " And A.����ʱ�� Between [1] And [2]"
    mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "�Һ�ʱ��"
    mcllFiter.Add Array(Trim(txtNOBegin.Text), Trim(txtNoEnd)), "�Һ�NO"
    mcllFiter.Add Array(Trim(txtFactBegin.Text), Trim(txtFactEnd)), "��Ʊ��"
    If cbo����Ա.ListIndex > 0 Then
        mcllFiter.Add NeedName(cbo����Ա.Text), "�Һ�Ա"
    Else
        mcllFiter.Add "", "�Һ�Ա"
    End If
    mcllFiter.Add "", "����"
    mcllFiter.Add "", "�����": mcllFiter.Add "", "���￨��"
    mcllFiter.Add "", "ҽ����": mcllFiter.Add "", "��������"
    mcllFiter.Add Val(IDKind.IDKind), "KIND"
    mcllFiter.Add "", "����ID"
    strKind = IDKind.GetCurCard.����

    mcllFiter.Add Trim(txtValue.Text), "_" & strKind
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    End If
    
    If cbo����Ա.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.����Ա����||''=[9]"
    If Trim(txtValue.Text) <> "" Then
        If mlngPrePatient <> 0 Then
            mstrFilter = mstrFilter & " And A.����ID=[12]"
            mcllFiter.Remove "����ID": mcllFiter.Add mlngPrePatient, "����ID"
        Else
            Select Case strKind
            Case "�����"
                mstrFilter = mstrFilter & " And A.����� = [11]"
                mcllFiter.Remove "�����": mcllFiter.Add Trim(txtValue.Text), "�����"
            Case "����", "��������￨"
                If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtValue.Text, 1))) > 0 Then
                    mstrFilter = mstrFilter & " And Upper(A.����) Like [8]"
                Else
                    mstrFilter = mstrFilter & " And A.���� Like [8]"
                End If
                mcllFiter.Remove "��������": mcllFiter.Add Trim(txtValue.Text), "��������"
            Case "ҽ����"
                mstrFilter = mstrFilter & " And B.ҽ����=[13]"
                mcllFiter.Remove "ҽ����": mcllFiter.Add Trim(txtValue.Text), "ҽ����"
            Case Else
                '��������,��ȡ��صĲ���ID
                '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
                '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
                '��7λ��,��ֻ��������,��Ȼȡ������
                lng�����ID = Val(IDKind.GetCurCard.�ӿ����)
                
                If lng�����ID <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, Trim(txtValue.Text), True, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(strKind, Trim(txtValue.Text), True, lng����ID, _
                        strPassWord, strErrMsg) = False Then lng����ID = 0
                End If
                If lng����ID = 0 Then
                    If strErrMsg = "" Then
                        MsgBox "δ�ҵ����������Ĳ���", vbInformation + vbOKOnly, gstrSysName
                        If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
                        zlControl.TxtSelAll txtValue
                        Exit Function
                    End If
                End If
                mstrFilter = mstrFilter & " And A.����ID=[12]"
                mcllFiter.Remove "����ID": mcllFiter.Add lng����ID, "����ID"
            End Select
        End If
    End If
    
    strSQL = ""
    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '�������Ʊ�ݺ��ж�,ֱ�Ӹ��ݵ��ݵķ���ʱ���ж�
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[5] ", " Between [5] And [6] ")
        strSQL = "Select A.NO" & _
        " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
        " Where A.��������=4 And A.ID=B.��ӡID And B.����=1" & _
        " And B.���� " & strSQLtmp
    End If
    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"
    '�Һſ���(ִ�п���)
    If cbo����.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And A.ִ�в���ID+0=[7]"
        mcllFiter.Remove "����"
        mcllFiter.Add cbo����.ItemData(cbo����.ListIndex), "����"
    End If
    mcllFiter.Add mstrFilter, "����"
    MakeFilter = True
End Function

Private Sub txtValue_Change()
    txtValue.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtValue Then
        'If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
       ' If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_GotFocus()
    Call zlControl.TxtSelAll(txtValue)
    Call zlCommFun.OpenIme(True)
    If txtValue.Text = "" And ActiveControl Is txtValue Then
'        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
'        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtvalue_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean

    On Error GoTo errH
    If txtValue.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub

    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
        End If
        txtValue.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    If blnCard And Len(txtValue.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtValue.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtValue.Text = txtValue.Text & Chr(KeyAscii)
            txtValue.SelStart = Len(txtValue.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtValue.Text)
    End If
    If Me.ActiveControl Is txtValue And mblnKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean

    On Error GoTo errH

    strSQL = ""
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.����ID=[2] "
        
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And B.�����=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]"
    Else
        Select Case objCard.����
        Case "����", "��������￨"
            txtValue.Tag = strInput
            Set mrsInfo = Nothing: Exit Sub
            zlCommFun.PressKey vbKeyTab
        Case "ҽ����"
            strInput = UCase(strInput)
            If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                '������ҽ������Ч:������:����:26982
                strSQL = strSQL & " And B.ҽ���� like [3] "
                strTemp = Left(strInput, 9) & "%"
            Else
                strSQL = strSQL & " And B.ҽ����=[1]"
            End If
        Case "���֤��", "���֤", "�������֤"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            strSQL = strSQL & " And B.����ID=[2]"
            strInput = "-" & lng����ID
        Case "IC����", "IC��"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            strSQL = strSQL & " And B.����ID=[2]"
            strInput = "-" & lng����ID
        Case "�����"
            If Not IsNumeric(strInput) Then strInput = "0"
            strSQL = strSQL & " And B.�����=[1]"
        Case Else
            '��������,��ȡ��صĲ���ID
            If Val(objCard.�ӿ����) > 0 Then
                lng�����ID = Val(objCard.�ӿ����)
                If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                If lng����ID = 0 Then lng����ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                                                        strPassWord, strErrMsg) = False Then lng����ID = 0
            End If
            If lng����ID <= 0 Then lng����ID = 0
            strSQL = strSQL & " And B.����ID=[2]"
            strInput = "-" & lng����ID
            blnHavePassWord = True
        End Select
    End If
    strSQL = "" & _
    "   Select distinct  B.����id As ID, Decode(sign(nvl(X.����id,0)),0,'','��') as �����˻�,  " & _
    "           B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ," & _
    "            A.���� ��������" & _
    "   From ������Ϣ B, ������� A,ҽ�ƿ���� Y,����ҽ�ƿ���Ϣ X" & _
    "   Where B.���� = A.���(+) and b.����id=X.����id(+)  " & _
    "               And X.״̬(+)=0 and  X.�����id=Y.id(+)  and Y.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   " & _
                    strSQL
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtValue.Hwnd)
    Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "��", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtValue.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%", dtpBegin.Value, dtpEnd.Value)
    
    If blnCancel Or mrsInfo Is Nothing Then
        Set mrsInfo = Nothing: txtValue.Text = "": Exit Sub
    End If
    
    If mrsInfo!id = 0 Then    'û���ҵ�������Ϣ
        Set mrsInfo = Nothing: txtValue.Text = "": Exit Sub
    End If
    
    txtValue.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtValue.Text = Nvl(mrsInfo!����)
    Me.txtValue.Tag = Nvl(mrsInfo!id)
    mlngPrePatient = Val(Nvl(mrsInfo!id))
    zlCommFun.PressKey vbKeyTab
    Exit Sub
    
NotFoundPati:
    Set mrsInfo = Nothing: txtValue.Text = "": Exit Sub
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
   '��ȡ������Ϣ
    Call GetPatient(objCard, txtValue.Text, blnCard)
End Sub

Private Sub InitCllData()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ����������
    '���ƣ����˺�
    '���ڣ�2010-06-02 15:44:19
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If mcllFiter Is Nothing Then
        Set mcllFiter = New Collection
        mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "�Һ�ʱ��"
        mcllFiter.Add Array("", ""), "�Һ�NO"
        mcllFiter.Add Array("", ""), "��Ʊ��"
        mcllFiter.Add "", "�Һ�Ա"
        mcllFiter.Add "", "����"
        mcllFiter.Add "", "�����": mcllFiter.Add "", "���￨��"
        mcllFiter.Add "", "ҽ����": mcllFiter.Add "", "��������"
        mcllFiter.Add 0, "KIND"
        mcllFiter.Add mstrFilter, "����"
        Exit Sub
    End If
    '�ָ�Ĭ������
    txtNOBegin.Text = mcllFiter("�Һ�NO")(0):    txtNoEnd.Text = mcllFiter("�Һ�NO")(1)
    txtFactBegin.Text = mcllFiter("��Ʊ��")(0):    txtFactEnd.Text = mcllFiter("��Ʊ��")(1)
    dtpBegin.Value = CDate(mcllFiter("�Һ�ʱ��")(0)):    dtpEnd.Value = CDate(mcllFiter("�Һ�ʱ��")(1))
    mstrFilter = CStr(mcllFiter("����"))

    '�����п��ܲ�����,���Բ�����ֵ
    Err = 0: On Error Resume Next
    If mcllFiter(Trim(IDKind.GetCurCard.����)) <> "" Then
        '��ʼ��
        txtValue.Text = mcllFiter("_" & Trim(IDKind.GetCurCard.����))
    End If
End Sub
