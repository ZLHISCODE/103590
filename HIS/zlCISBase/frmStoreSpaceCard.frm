VERSION 5.00
Begin VB.Form frmStoreSpaceCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��λ�༭"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   Icon            =   "frmStoreSpaceCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboRoom 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   2640
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9810
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   2625
   End
   Begin VB.TextBox txt��ע 
      Appearance      =   0  'Flat
      Height          =   1020
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   2625
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   2625
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      Top             =   3480
      Width           =   1100
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "������������0"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   3570
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   480
      TabIndex        =   11
      Top             =   780
      Width           =   360
   End
   Begin VB.Label lbl��ע 
      AutoSize        =   -1  'True
      Caption         =   "��ע"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      Top             =   2220
      Width           =   360
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblSpace 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1260
      Width           =   360
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   180
      Width           =   360
   End
End
Attribute VB_Name = "frmStoreSpaceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭״̬ As Integer '1-���� 2-�޸�
Private mlng�ⷿID As Long
Private mlng��λid As Long
Private mblnRefresh As Boolean
Private mintAddCount As Integer '��������


Private Function GetNextCode() As String
    Dim rsTemp As ADODB.Recordset
    
    'ȡ��һ������
    gstrSql = "Select Max(����) as ���� From ҩƷ�ⷿ��λ "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "�ⷿ��λ������")
    
    If NVL(rsTemp!����) = "" Then
        GetNextCode = "00001"
    Else
        GetNextCode = zlCommFun.IncStr(rsTemp!����)
    End If
End Function

Public Function ShowMe(ByVal int�༭״̬ As Integer, ByVal lng�ⷿID As Long, ByVal lng��λid As Long, ByVal fraPar As Form) As Boolean
    
    mint�༭״̬ = int�༭״̬
    mlng�ⷿID = lng�ⷿID
    mlng��λid = lng��λid
    
    Me.Show vbModal, fraPar
    
    ShowMe = mblnRefresh
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim colData As New Collection
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    
    If Trim(txt����.Text) = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt����.Text, vbFromUnicode)) > txt����.MaxLength Then
        MsgBox "���Ƴ��ȳ���" & txt����.MaxLength & "���ַ���", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If Trim(txt����.Text) = "" Then
        MsgBox "��λ���Ʋ���Ϊ�գ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt����.Text, vbFromUnicode)) > txt����.MaxLength Then
        MsgBox "���Ƴ��ȳ���" & txt����.MaxLength & "���ַ����� " & Int(txt����.MaxLength / 2) & "�����֣�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If Trim(txt����.Text) = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt����.Text, vbFromUnicode)) > txt����.MaxLength Then
        MsgBox "���볤�ȳ���" & txt����.MaxLength & "���ַ�����" & Int(txt����.MaxLength / 2) & "�����֣�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt��ע.Text, vbFromUnicode)) > txt��ע.MaxLength Then
        MsgBox "��ע���ȳ���" & txt��ע.MaxLength & "���ַ�����" & Int(txt��ע.MaxLength / 2) & "�����֣�", vbInformation, gstrSysName
        txt��ע.SetFocus
        Exit Sub
    End If
    
    '�������ظ�
    If mint�༭״̬ = 1 Then
        '����ʱȫ����
        gstrSql = "Select 1 From ҩƷ�ⷿ��λ Where ���� = [1]"
    Else
        '�޸�ʱ���ų�����
        gstrSql = "Select 1 From ҩƷ�ⷿ��λ Where ���� = [1] And ID <> [2] "
    End If
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "�������ظ�", txt����.Text, mlng��λid)
    
    If Not rsData.EOF Then
        MsgBox "�����ظ���������¼�룡", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    '��������ظ�
    If mint�༭״̬ = 1 Then
        '����ʱȫ����
        gstrSql = "Select 1 From ҩƷ�ⷿ��λ Where �ⷿid = [2] And ���� = [1]"
    Else
        '�޸�ʱ���ų�����
        gstrSql = "Select 1 From ҩƷ�ⷿ��λ Where �ⷿid = [2] And ���� = [1] And ID <> [3] "
    End If
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "��������ظ�", txt����.Text, Val(cboRoom.ItemData(cboRoom.ListIndex)), mlng��λid)
    
    If Not rsData.EOF Then
        MsgBox "�����ظ���������¼�룡", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    
    If mint�༭״̬ = 1 Then
        '����
        gstrSql = "Zl_ҩƷ�ⷿ��λ_Insert("
        '����
        gstrSql = gstrSql & "'" & txt����.Text & "'"
        '����_In   In ҩƷ�ⷿ��λ.����%Type,
        gstrSql = gstrSql & ",'" & txt����.Text & "'"
        '����_In   In ҩƷ�ⷿ��λ.����%Type,
        gstrSql = gstrSql & ",'" & txt����.Text & "'"
        '�ⷿid_In In ҩƷ�ⷿ��λ.�ⷿid%Type
        gstrSql = gstrSql & "," & Val(cboRoom.ItemData(cboRoom.ListIndex))
        '��ע_In In ҩƷ�ⷿ��λ.��ע%Type
        gstrSql = gstrSql & "," & IIf(txt��ע.Text = "", "null", "'" & txt��ע.Text & "'")
        gstrSql = gstrSql & ")"
        
        colData.Add gstrSql, "k_1"
    Else
        '�޸�
        gstrSql = "Zl_ҩƷ�ⷿ��λ_Update("
        'ID
        gstrSql = gstrSql & mlng��λid
        '����
        gstrSql = gstrSql & ",'" & txt����.Text & "'"
        '����_In   In ҩƷ�ⷿ��λ.����%Type,
        gstrSql = gstrSql & ",'" & txt����.Text & "'"
        '����_In   In ҩƷ�ⷿ��λ.����%Type,
        gstrSql = gstrSql & ",'" & txt����.Text & "'"
        '�ⷿid_In In ҩƷ�ⷿ��λ.�ⷿid%Type
        gstrSql = gstrSql & "," & Val(cboRoom.ItemData(cboRoom.ListIndex))
        '��ע_In In ҩƷ�ⷿ��λ.��ע%Type
        gstrSql = gstrSql & "," & IIf(txt��ע.Text = "", "null", "'" & txt��ע.Text & "'")
        gstrSql = gstrSql & ")"
        
        colData.Add gstrSql, "k_2"
    End If
    
    gcnOracle.BeginTrans
    For i = 1 To colData.Count
        Call zldatabase.ExecuteProcedure(colData(i), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    mblnRefresh = colData.Count > 0
    
    If mint�༭״̬ = 1 Then
        '��ս������ݣ�������������
        txt����.Text = GetNextCode
        txt����.Text = ""
        txt����.Text = ""
        txt��ע.Text = ""
        txt����.SetFocus
        
        mintAddCount = mintAddCount + 1
        If lblComment.Visible = False Then lblComment.Visible = True
        lblComment.Caption = "������������" & mintAddCount
    Else
        Unload Me
    End If
    
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    txt����.SetFocus
    
    lblComment.Visible = (mint�༭״̬ = 1)
    lblComment.Caption = "������������0"
End Sub

Private Sub SetBorder(ByVal objControl As Variant, Optional ByVal blnIsFocuse As Boolean = True)
    '���ܣ������ı��򱳾�ɫ
    
    If blnIsFocuse Then
        objControl.BackColor = &HDCDBC5
    Else
        objControl.BackColor = &H80000005
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    mintAddCount = 0
    
    With cboRoom
        .Clear
        For i = 0 To frmStoreSpace.cboRoom.ListCount - 1
            .AddItem frmStoreSpace.cboRoom.List(i)
            .ItemData(.NewIndex) = frmStoreSpace.cboRoom.ItemData(i)
        Next
        
        If .ListIndex <> 0 Then .ListIndex = frmStoreSpace.cboRoom.ListIndex
        
        .Enabled = False
    End With
    
    gstrSql = "Select ID, ����, ����, ����, �ⷿid, ��ע From ҩƷ�ⷿ��λ where id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "�ⷿ��λ", mlng��λid)
                
    txt����.MaxLength = rsTemp.Fields("����").DefinedSize
    txt����.MaxLength = rsTemp.Fields("����").DefinedSize
    txt����.MaxLength = rsTemp.Fields("����").DefinedSize
    txt��ע.MaxLength = rsTemp.Fields("��ע").DefinedSize
        
    If Not rsTemp.EOF Then
        txt����.Text = rsTemp!����
        txt����.Text = rsTemp!����
        txt����.Text = NVL(rsTemp!����)
        txt��ע.Text = NVL(rsTemp!��ע)
    End If
    
    If mint�༭״̬ = 1 Then
        txt����.Text = GetNextCode
    End If
End Sub


Private Sub txt��ע_GotFocus()
    zlControl.TxtSelAll txt��ע
    SetBorder txt��ע
End Sub

Private Sub txt��ע_LostFocus()
    SetBorder txt��ע, False
End Sub


Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    SetBorder txt����
End Sub

Private Sub txt����_LostFocus()
    SetBorder txt����, False
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    txt����.Text = String(txt����.MaxLength - Len(txt����.Text), "0") & txt����.Text
End Sub

Private Sub txt����_LostFocus()
    SetBorder txt����, False
End Sub

Private Sub txt����_Change()
    Dim strTmp As String
    
    strTmp = MoveSpecialChar(txt����.Text)
    If txt����.Text <> strTmp Then
        txt����.Text = strTmp
    End If
    Me.txt����.Text = zlStr.GetCodeByORCL(strTmp, False, 10)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    SetBorder txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        If LenB(StrConv(txt����.Text, vbFromUnicode)) >= 50 Then KeyAscii = 0
    End If
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        If LenB(StrConv(txt��ע.Text, vbFromUnicode)) >= 100 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    SetBorder txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Not (KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        If LenB(StrConv(txt����.Text, vbFromUnicode)) >= 10 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
End Sub

Private Sub txt����_LostFocus()
    Me.txt����.Text = zlStr.GetCodeByORCL(txt����.Text, False, 10)
    SetBorder txt����, False
End Sub


