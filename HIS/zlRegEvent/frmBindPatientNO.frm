VERSION 5.00
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmBindPatientNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ű�"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmBindPatientNO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4815
   StartUpPosition =   1  '����������
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Appearance      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   10.5
      FontName        =   "����"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txt����� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   15
      Top             =   735
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   4905
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      TabIndex        =   3
      ToolTipText     =   "�ȼ�:F11"
      Top             =   240
      Width           =   2400
   End
   Begin VB.CommandButton cmdYb 
      Caption         =   "ҽ��"
      Height          =   345
      Left            =   4020
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      ToolTipText     =   "�ȼ���F2"
      Top             =   2880
      Width           =   1110
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2385
      TabIndex        =   0
      Top             =   2880
      Width           =   1110
   End
   Begin VB.Label txt���֤ 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   13
      Top             =   1740
      Width           =   3600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "���֤"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2835
      TabIndex        =   11
      Top             =   2250
      Width           =   420
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   450
      TabIndex        =   10
      Top             =   1290
      Width           =   420
   End
   Begin VB.Label txt���� 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   9
      Top             =   1230
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   450
      TabIndex        =   8
      Top             =   2250
      Width           =   420
   End
   Begin VB.Label txt�Ա� 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   2190
      Width           =   1200
   End
   Begin VB.Label txt���� 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3345
      TabIndex        =   6
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   810
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   450
      TabIndex        =   4
      Top             =   300
      Width           =   420
   End
End
Attribute VB_Name = "frmBindPatientNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mintԤԼʧЧ���� As Integer
Private mstrPrivs As String, mintIDKind As Integer
Private mstrNo As String
Private mblnOk As Boolean
Private mbln����סԺ���˹Һ� As Boolean
Private mblnNotClick As Boolean
Private mstrYBPati   As String
Private mintInsure   As Integer
Private mlngPatient As Long
Private Const mlngModule = 1111
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String

Private Sub cmdEnter_Click()
       Dim blnNoPrint   As Boolean
       If mrsInfo Is Nothing Then Exit Sub
       If mrsInfo.RecordCount = 0 Then Exit Sub
       If txtPatient.Text = "" Or mrsInfo!����ID <> txtPatient.Tag Then Exit Sub
       cmdEnter.Enabled = False
       If SaveData() = False Then
            cmdEnter.Enabled = True
            Exit Sub
       End If
       Select Case gByt��ӡ��������
       Case 0: blnNoPrint = True
       Case 1: blnNoPrint = False
       Case 2:
              If MsgBox("�Ƿ���Ҫ��ӡ�������룿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnNoPrint = True
              End If
       End Select
       If Not blnNoPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me, "����ID=" & Val(Nvl(mrsInfo!����ID)), 2)
       End If
       
       Call ClearControlContent
End Sub

Private Sub ClearControlContent()
        Me.txtPatient.Tag = ""
        mlngPatient = 0
        Me.txt�����.Tag = ""
        txt����.Caption = ""
        txt�Ա�.Caption = ""
        txt����.Caption = ""
        txtPatient.Text = ""
        Me.txt�����.Text = ""
        txt���֤.Caption = ""
        Set mrsInfo = Nothing
       cmdEnter.Enabled = True
End Sub
'-----------------------------------------------------------------------------------
Private Sub cmdYb_Click()
     'ҽ�����֤��֤
     Call zlInusreIdentify
     
End Sub
Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ҽ������鿨
    '���ƣ����˺�
    '���ڣ�2010-07-14 11:32:08
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim str�������� As String
    If mrsInfo Is Nothing Then
        lng����ID = 0
        str�������� = ""
    Else
        lng����ID = Val(Nvl(mrsInfo!����ID))
        str�������� = Nvl(mrsInfo!��������)
    End If
     
    mstrYBPati = gclsInsure.Identify(3, lng����ID, mintInsure)
    If mstrYBPati <> "" Then
        '�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
        If UBound(Split(mstrYBPati, ";")) >= 8 Then
            If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng����ID = Val(Split(mstrYBPati, ";")(8))
        End If
        If lng����ID <> 0 Then
            '����:29283
            '  -- ����:���ó���-1-�Һ�;2-�շ�
            '  --        ����id_In-����ID(δ������,������)
            '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
            '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
            txtPatient.Text = "-" & lng����ID
            txtPatient_KeyPress (13)
            If str�������� = "" Then txtPatient.ForeColor = vbRed
            Me.txtPatient.SetFocus
        Else
            mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        End If
    Else
        '�޸����⣺38917 ���ߣ�Ƚ��
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
    End If
End Sub

 
Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
    
End Sub

Private Sub cmdOK_Click()
     If Me.txtPatient.Text = "" Or Me.txtPatient.Text <> Me.txtPatient.Tag Then Exit Sub
        
End Sub

Private Function CheckPatient(ByVal str����� As String, ByVal lng����ID As Long) As Boolean

'���ܣ��ж�ָ��������Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strMsg As String
    Dim lng�ϲ�ID As Long
    On Error GoTo errH
    strSQL = "    " & vbNewLine & " Select B.����id As ID, B.����id, B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
    strSQL = strSQL & vbNewLine & "      A.���� ��������"
    strSQL = strSQL & vbNewLine & " From ������Ϣ B, ������� A"
    strSQL = strSQL & vbNewLine & " Where B.���� = A.���(+) And B.ͣ��ʱ�� Is Null  "
    strSQL = strSQL & vbNewLine & " And b.�����=[1] And ����ID<>[2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str�����, lng����ID)
    If rsTmp.RecordCount > 0 Then
       If Nvl(mrsInfo!����) <> Nvl(rsTmp!����) Then
            strMsg = "��ǰ��������[" & Nvl(mrsInfo!����) & "]�����������[" & str����� & "]�Ĳ���[" & Nvl(rsTmp!����) & "]�Ĳ��˲�һ��!" & vbCrLf & _
                    "�Ƿ������Ϊ[" & Nvl(mrsInfo!����) & "]����Ϣ�ϲ�������Ϊ[" & rsTmp!���� & "]�У�"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            lng�ϲ�ID = Val(rsTmp!����ID)
            If zlPatiMerge(lng����ID, lng�ϲ�ID, False) = False Then Exit Function
       Else
            strMsg = "��ǰϵͳ�Ѿ���������Ϊ[" & mrsInfo!���� & "]���������Ϊ[" & str����� & "]�Ĳ���,�Ƿ񽫵�ǰ���˺ϲ��������Ϊ[" & str����� & "]�Ĳ�����?"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            lng�ϲ�ID = Val(rsTmp!����ID)
            If zlPatiMerge(lng����ID, lng�ϲ�ID, False) = False Then Exit Function
            Call GetPatient("-" & lng�ϲ�ID)
       End If
        strSQL = "    " & vbNewLine & " Select B.����id As ID, B.����id, B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
        strSQL = strSQL & vbNewLine & "      A.���� ��������"
        strSQL = strSQL & vbNewLine & " From ������Ϣ B, ������� A"
        strSQL = strSQL & vbNewLine & " Where B.���� = A.���(+) And B.ͣ��ʱ�� Is Null  "
        strSQL = strSQL & vbNewLine & " And b.����ID=[1] "
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ϲ�ID)
        mlngPatient = lng�ϲ�ID
    End If
      CheckPatient = True
       Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function SaveData() As Boolean
    '
    Dim strSQL     As String
    Dim lng����ID  As Long
    Dim strҽ��    As String
    Dim str�����  As String
    Dim strDat     As String
    Dim Datsys     As Date
    Dim intInsure  As Integer
    On Error GoTo Hd
    
    If txt�����.Text = "" Then
        MsgBox "����Ų���Ϊ��!", vbInformation, Me.Caption
        Exit Function
    End If
    str����� = txt�����.Text
    If CheckPatient(str�����, mlngPatient) = False Then Exit Function
    If Exist�����(Me.txt�����.Text, mlngPatient) Then
        MsgBox "������Ѿ���ʹ��!", vbInformation, Me.Caption
        Exit Function
    End If
    Datsys = zlDatabase.Currentdate
    strDat = "to_date('" & Format(Datsys, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    If Trim(Me.txt����.Caption) <> "" And mstrYBPati <> "" Then
        strҽ�� = Split(mstrYBPati, ";")(1)
        intInsure = mintInsure
    End If
  'Zl_������Ϣ_�������(
  '����id_In   ������Ϣ.����id%Type,
  '�����_In   ������Ϣ.�����%Type,
  '�Ǽ�ʱ��_In ������Ϣ.�Ǽ�ʱ��%Type,
  'ҽ����_In   ������Ϣ.ҽ����%Type := Null,
  '����_In     ������Ϣ.����%Type := Null,
  '��������_In Number:=0
  '--���ܣ�����ҺŲ�����Ϣ����� ��
  '--������
  '--�������ͣ�
  '--             0=�����������  ���Բ��˵ķ��ü�¼ �͹Һż�¼ ����� ���и���
  '--             1=���������  ͬʱ���� ���˵� ������Ϣ
    strSQL = "Zl_������Ϣ_�������"
    strSQL = strSQL & "(" & mlngPatient & ","
    strSQL = strSQL & "'" & str����� & "'" & ","
    strSQL = strSQL & strDat & ","
    strSQL = strSQL & IIf(strҽ�� = "", "NUll", "'" & strҽ�� & "'") & ","
    strSQL = strSQL & IIf(intInsure = 0, "NULL", intInsure) & ","
    strSQL = strSQL & "1)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    SaveData = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    Select Case KeyCode
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKind.Enabled Then IDKind.IDKind = IDKind.GetKindIndex("IC����"): Call IDKind_Click(IDKind.GetCurCard)
            ElseIf Me.ActiveControl Is txtPatient Then
                If IDKind.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDkindStr, ";")), IDKind.IDKind - 1)
                    Else
                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), 0, IDKind.IDKind + 1)
                    End If
                End If
            End If
        Case vbKeyF11
            If txtPatient.Enabled And txtPatient.Visible And Not txtPatient.Locked Then
                If Me.ActiveControl Is txtPatient Then
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), IDKind.GetKindIndex("����"), IDKind.IDKind + 1)
                Else
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyReturn
       
    End Select
End Sub
Private Sub Form_Load()
    Dim strTemp As String
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.Hwnd)
    Call mobjICCard.SetParent(Me.Hwnd)
    If gobjSquare.objSquareCard Is Nothing Then
        CreateSquareCardObject gfrmMain, mlngModule
    End If
    InitIDKind
    Set mobjICCard.gcnOracle = gcnOracle
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mbln����סԺ���˹Һ� = zlDatabase.GetPara("����סԺ���˹Һ�", glngSys, mlngModule, 0) = "1"
    Me.Icon = frmRegist.Icon
End Sub

Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModule, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "��������") > 0
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
         
       
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    
        
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Set mrsInfo = Nothing
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
     
End Sub


Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If IsCardType(IDKind, "IC����") Then
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(Trim(txtPatient))
            End If
        End If
        Exit Sub
    End If
    lng�����ID = IDKind.GetCurCard.�ӿ����
    If lng�����ID = 0 Then Exit Sub
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
    Call GetPatient(Trim(txtPatient.Text))
    End If
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
     '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
     
    Set gobjSquare.objCurCard = objCard
     
    If objCard.�ӿ���� > 0 Then
        txtPatient.MaxLength = IDKind.GetCardNoLen
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    Else
        txtPatient.MaxLength = 0: txtPatient.PasswordChar = ""
    End If
    
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_KeyPress(KeyAscii As Integer)
'
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

    txtPatient.Text = objPatiInfor.����
    Call txtPatient_KeyPress(vbKeyReturn)
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub
Private Sub txtPatient_Change()
    txtPatient.Tag = ""
  
    txtPatient.ForeColor = &H80000008
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
   
End Sub

Private Sub txtPatient_GotFocus()
   zlControl.TxtSelAll txtPatient
      If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
  Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IsCardType(IDKind, "����") Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys)
    ElseIf IsCardType(IDKind, "�����") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
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
    Call SaveErrLog
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
    Dim blnIDCard As Boolean
    On Error GoTo errH
    If Not mbln����סԺ���˹Һ� Then
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=B.����ID And ��ҳID=B.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If
    
    strSQL = ""
    
    If Not (blnCard Or IDKind.GetCurCard.�ӿ���� = IDKind.GetfaultCard.�ӿ����) _
         And IDKind.GetCurCard.���� Like "*����" And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
      
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
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
    Else
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                '����
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!���� Then blnSame = True
                End If
                 
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(txtPatient.Text) < 2) Then
                        txt����.Caption = "": txt�Ա� = "": txt���� = ""
                        txtPatient.Text = "": Me.txt�����.Text = ""
                        txt����.Caption = "0"
                        txt���֤.Caption = ""
                        txt�Ա�.Caption = ""
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
'                 strInput = UCase(strInput)
'                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
'                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
'                strInput = "-" & lng����ID
                 blnIDCard = True
                 strSQL = strSQL & " And B.���֤��=[1] " & str����Ժ
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
                strInput = "-" & lng����ID
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.�����=[1]" & str����Ժ
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
    strSQL = "    " & vbNewLine & " Select /*+Rule */distinct  B.����id As ID, Decode(sign(nvl(ylkxx.����id,0)),0,'','��') as �����˻�, B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ,"
    strSQL = strSQL & vbNewLine & "      A.���� ��������,B.��������"
    strSQL = strSQL & vbNewLine & " From ������Ϣ B, ������� A,ҽ�ƿ���� YLK,����ҽ�ƿ���Ϣ YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.���� = A.���(+) and b.����id=ylkxx.����id(+) and ylkxx.״̬(+)=0 and  ylkxx.�����id=ylk.id(+)  and ylk.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
     
    On Error GoTo errH
     If Not blnIDCard Then
        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strInput & "%")
     Else
        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
     Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "��", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%")
     End If
     If Not mrsInfo Is Nothing And Not blnCancel Then  'And Not blnCancel
        If mrsInfo.RecordCount = 0 Then
            Set mrsInfo = Nothing
            txt����.Caption = "": txt�Ա� = "": txt���� = ""
            txtPatient.Text = "": Me.txt�����.Text = ""
            txt����.Caption = "0"
            txt���֤.Caption = ""
            txt�Ա�.Caption = ""
            Exit Sub
        ElseIf mrsInfo!ID = 0 Then  'û���ҵ�������Ϣ
            Set mrsInfo = Nothing
            txt����.Caption = "": txt�Ա� = "": txt���� = ""
            txtPatient.Text = "": Me.txt�����.Text = ""
            txt����.Caption = "0"
            txt���֤.Caption = ""
            txt�Ա�.Caption = ""
            Exit Sub
        Else '��ȡ��������Ϣ
          
            Me.txt�����.Tag = Nvl(mrsInfo!�����)
            txt����.Caption = Nvl(mrsInfo!��������)
            txt�Ա�.Caption = Nvl(mrsInfo!�Ա�)
            txt����.Caption = Nvl(mrsInfo!����)
            txtPatient.Text = Nvl(mrsInfo!����)
            Me.txtPatient.Tag = Nvl(mrsInfo!ID)
            Me.txt�����.Text = Nvl(mrsInfo!�����)
            txt���֤.Caption = Nvl(mrsInfo!���֤��)
            mlngPatient = Val(Nvl(mrsInfo!ID))
            '74428:���ϴ���2014-7-7������������ɫ����
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(Trim(txt����.Caption) <> "", vbRed, txtPatient.ForeColor))
        End If
    Else 'ȡ��ѡ��
        txt����.Caption = "": txt�Ա� = "": txt���� = ""
        txtPatient.Text = "": Me.txt�����.Text = ""
        mlngPatient = 0
        txt����.Caption = "0"
        txt���֤.Caption = ""
        txt�Ա�.Caption = ""
        Set mrsInfo = Nothing: Exit Sub
    End If
    
'    ''''''''''''''''
'
'     If Not mrsInfo Is Nothing And Not blnCancel Then
'        If mrsInfo!ID = 0 Then 'û���ҵ�������Ϣ
'            Set mrsInfo = Nothing
'            txt����.Caption = "": txt�Ա� = "": txt���� = ""
'            txtPatient.Text = "": Me.txt�����.Text = ""
'            txt����.Caption = "0"
'            txt���֤.Caption = ""
'            txt�Ա�.Caption = ""
'            Exit Sub
'        Else '��ȡ��������Ϣ
'
'          Me.txt�����.Tag = Nvl(mrsInfo!�����)
'          txt����.Caption = Nvl(mrsInfo!��������)
'          txt�Ա�.Caption = Nvl(mrsInfo!�Ա�)
'          txt����.Caption = Nvl(mrsInfo!����)
'          txtPatient.Text = Nvl(mrsInfo!����)
'          Me.txtPatient.Tag = Nvl(mrsInfo!ID)
'          Me.txt�����.Text = Nvl(mrsInfo!�����)
'          txt���֤.Caption = Nvl(mrsInfo!���֤��)
'          mlngPatient = Val(Nvl(mrsInfo!ID))
'          If Trim(txt����.Caption) <> "" Then Me.txtPatient.ForeColor = vbRed
'        End If
'    Else 'ȡ��ѡ��
'        txt����.Caption = "": txt�Ա� = "": txt���� = ""
'        txtPatient.Text = "": Me.txt�����.Text = ""
'        mlngPatient = 0
'        txt����.Caption = "0"
'        txt���֤.Caption = ""
'        txt�Ա�.Caption = ""
'        Set mrsInfo = Nothing: Exit Sub
'    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
 
 

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        Call GetPatient(Trim(txtPatient.Text))
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            Call GetPatient(Trim(txtPatient.Text))
        Else
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then mobjICCard.SetEnabled (txtPatient.Text = "")
    End If
End Sub


Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
End Sub

Private Sub txt�����_Change()
   Me.txt�����.Tag = ""
End Sub

Private Sub txt�����_GotFocus()
   zlControl.TxtSelAll txt�����
End Sub

 
'Private Function CheckPatiValid(ByVal strCard As String) As Boolean
'    '------------------------------------------------------------------------------------------------------------------------
'    '���ܣ����ָ������Ŀ����Ƿ�Ϸ�
'    '��Σ�strCard-ָ���Ŀ���
'    '���أ��Ϸ�,����True,���򷵻�False
'    '���ƣ����˺�
'    '���ڣ�2010-07-19 10:14:31
'    '˵����31182
'    '------------------------------------------------------------------------------------------------------------------------
'   Dim rsTmp As ADODB.Recordset, strSQL As String, lng����ID As Long
'
'    strSQL = "Select Nvl(����״̬,0) ����״̬,����ID,����,�Ա� From ������Ϣ Where ���￨�� = [1]"
'    On Error GoTo errH
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCard)
'    If rsTmp.RecordCount = 0 Then CheckPatiValid = True: Exit Function
'
'    '1.���״̬:ԭ����Ҫ��������￨ʱ���м���,����txt����_Validate����,��һ���ܼ�鵽,���,�������ڰ�ȷ��ʱ,���Ӹü��
'    If Val(Nvl(rsTmp!����״̬)) <> 0 Then
'        MsgBox "����Ϊ" & strCard & "�Ĳ������ھ����ȴ�����,���ܰ󶨸ÿ���.", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '2.����Ƿ���������ͬ
'    If Nvl(rsTmp!����) <> Trim(txtPatient.Text) And Val(txt����.Tag) = 0 Then
'       If MsgBox("�ֿ����ˡ�" & Nvl(rsTmp!����) & "��������Ĳ��ˡ�" & Trim(txtPatient.Text) & "����һ��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
'    End If
'
'    '3.�ҺŲ�����ˢ���￨�ó��Ĳ�����������ͬ�����Ĳ���
'    lng����ID = Val(Nvl(rsTmp!����ID))
'    If Val(txt����.Tag) <> lng����ID And Val(txt����.Tag) <> 0 Then
'        If Nvl(rsTmp!����) <> Trim(txtPatient.Text) Then
'            If MsgBox("ע��: " & vbCrLf & _
'                             "     �ֿ����ˡ�" & Nvl(rsTmp!����) & "��������Ĳ��ˡ�" & Trim(txtPatient.Text) & "����һ��," & vbCrLf & _
'                             "     ��ͬʱ���ǽ�������,�Ƿ񽫲��ˡ�" & Trim(txtPatient.Text) & "���ϲ������ˡ�" & Nvl(rsTmp!����) & "����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
'            '�ϲ�
'            If zlPatiMerge(Val(txt����.Tag), lng����ID, True) = False Then Exit Function
'        Else '����������ͬ,�Զ����кϲ�
'            '�Զ��ϲ�
'            If zlPatiMerge(Val(txt����.Tag), lng����ID, False) = False Then Exit Function
'        End If
'        '����ˢ����ص�����
'        RaiseEvent PatiMerged(lng����ID)
'
'    End If
'    CheckPatiValid = True
'    Exit Function
'errH:
'    If errCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Private Sub txt�����_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
End Sub


'��ȡidkind��Ĭ��kindֵ
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
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
     End Select
End Function
                

