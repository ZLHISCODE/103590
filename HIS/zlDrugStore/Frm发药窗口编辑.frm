VERSION 5.00
Begin VB.Form Frm��ҩ���ڱ༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ҩ���ڱ༭"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "Frm��ҩ���ڱ༭.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Cmd���� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3990
      TabIndex        =   4
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   5
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   150
      TabIndex        =   6
      Top             =   0
      Width           =   3675
      Begin VB.ComboBox cboWindow 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1500
         Width           =   2085
      End
      Begin VB.CheckBox Chkר�� 
         Caption         =   "ר��"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1020
         MaxLength       =   1
         TabIndex        =   0
         Top             =   270
         Width           =   500
      End
      Begin VB.CommandButton Cmdҩ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   2820
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1050
         Width           =   285
      End
      Begin VB.TextBox Txtҩ�� 
         Height          =   300
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   1
         Top             =   660
         Width           =   2085
      End
      Begin VB.Label lblWindow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�кŴ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Lblҩ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   8
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
   End
End
Attribute VB_Name = "Frm��ҩ���ڱ༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IntEditState As Integer '1:����;2:�޸�
Public gStr���� As String
Public gLngҩ��ID As Long
Private mstr���� As String




Private Sub Cmd����_Click()
    Dim RecCheck As New ADODB.Recordset
    If CheckData = False Then Exit Sub
    
    
    On Error GoTo ErrHand
    If EditState = 1 Then
        gstrSQL = "Select Count(*) Records From ��ҩ���� Where ҩ��ID=[1] And ����=[2] "
        Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[������]", Me.Txtҩ��.Tag, Me.Txt����)
        
        With RecCheck
            If Not .EOF Then
                If !Records <> 0 Then
                    MsgBox "��ҩ���ķ�ҩ����[" & Txt���� & "]�Ѵ��ڣ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End With
    End If
    
    If mstr���� <> "" And mstr���� <> Txt����.Text Then
        gstrSQL = " zl_��ҩ����_ҵ����� (" & Me.Txtҩ��.Tag & ",'" & mstr���� & "','" & Me.Txt����.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���·�ҩ����")
    End If
    
    gcnOracle.BeginTrans
    If EditState = 1 Then
        gstrSQL = " zl_��ҩ����_insert ('" & Txt���� & "','" & Txt���� & "',1," & Me.Txtҩ��.Tag & "," & Chkר��.Value & ",'" & Trim(Me.cboWindow.Text) & "')"
    Else
        gstrSQL = " zl_��ҩ����_update ('" & Txt���� & "','" & Txt���� & "'," & Me.Txtҩ��.Tag & "," & Chkר��.Value & ",'" & gStr���� & "'," & gLngҩ��ID & ",'" & Trim(Me.cboWindow.Text) & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-���·�ҩ����")
    gcnOracle.CommitTrans
    
    If EditState = 2 Then
        Unload Me
        Exit Sub
    End If
    
    Me.Txt���� = GetMaxCode(Txtҩ��.Tag)
    Me.Txt���� = ""
    Me.Txt����.SetFocus
    mstr���� = ""
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Public Property Get EditState() As Variant
    EditState = IntEditState
End Property

Public Property Let EditState(ByVal vNewValue As Variant)
    IntEditState = vNewValue
End Property

Private Function CheckData() As Boolean
    CheckData = False
    
    If Txt���� = "" Then
        MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
        Me.Txt����.SetFocus
        Exit Function
    End If
    If Not IsNumeric(Txt����) Then
        MsgBox "����Ӧ��Ϊ�����ͣ�", vbInformation, gstrSysName
        Me.Txt����.SetFocus
        Exit Function
    End If
    If Txt���� = "" Then
        MsgBox "���Ʋ���Ϊ�գ�", vbInformation, gstrSysName
        Me.Txt����.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Txt����, vbFromUnicode)) > 10 Then
        MsgBox "���Ƴ��ȳ��������5�����ֻ�10���ַ�����", vbInformation, gstrSysName
        Me.Txt����.SetFocus
        Exit Function
    End If
    If Val(Me.Txtҩ��.Tag) = 0 Then
        MsgBox "��ѡ��ҩ����", vbInformation, gstrSysName
        Me.Txtҩ��.SetFocus
        Exit Function
    End If
    
    CheckData = True
End Function

Private Sub Cmdҩ��_Click()
    Dim RecTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select ID,����,���� From ���ű� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (" & _
          " Select distinct ����ID From ��������˵��" & _
          " Where �������� Like '%ҩ��')" & _
          " And To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' order by ����"
    Call zlDatabase.OpenRecordset(RecTmp, gstrSQL, "��ȡ����ҩ��")
    
    With FrmNodeSelect
        Set .TreeRec = RecTmp.Clone
        .StrNode = "����ҩ��"
        .Show 1, Me
        If .BlnSuccess Then
            Me.Txtҩ�� = .CurrentName
            Me.Txtҩ��.Tag = .CurrentID
        Else
            Me.Txtҩ�� = ""
            Me.Txtҩ��.Tag = 0
        End If
        Unload FrmNodeSelect
    End With
    
    Call LoadWindow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    If EditState = 1 Then
        If frm��ҩ����.Tree.SelectedItem.Key = "R" Then
            Txtҩ�� = ""
            Txtҩ��.Tag = 0
        Else
            Me.Txtҩ�� = Mid(frm��ҩ����.Tree.SelectedItem, InStr(1, frm��ҩ����.Tree.SelectedItem, "��") + 1)
            Me.Txtҩ��.Tag = Mid(frm��ҩ����.Tree.SelectedItem.Key, 3)
            Call LoadWindow
        End If
        Txt���� = GetMaxCode(Txtҩ��.Tag)
        Exit Sub
    End If
    
    Me.Txt���� = frm��ҩ����.Lvw.SelectedItem.SubItems(1)
    Me.Txt���� = frm��ҩ����.Lvw.SelectedItem
    Me.Chkר�� = IIf(frm��ҩ����.Lvw.SelectedItem.SubItems(4) = "��", 1, 0)
    Me.Txtҩ�� = frm��ҩ����.Lvw.SelectedItem.SubItems(3)
    Me.Txtҩ��.Tag = Mid(frm��ҩ����.Lvw.SelectedItem.Key, 3, InStr(1, frm��ҩ����.Lvw.SelectedItem.Key, ",") - 3)
    mstr���� = Txt����.Text
    
    gStr���� = Txt����
    gLngҩ��ID = Me.Txtҩ��.Tag
    
    Call LoadWindow
    
End Sub

Private Sub LoadWindow()
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim intRow As Integer
    
    On Error GoTo errHandle
    
    strSQL = "select ���� from ��ҩ���� where ҩ��id=[1] and �кŴ��� is null and ����<>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���д���", Txtҩ��.Tag, IIf(Me.Txt����.Text = "", " ", Me.Txt����.Text))
    
    intRow = -1
    Me.cboWindow.Clear
    Me.cboWindow.AddItem " "
    Do While Not rsTemp.EOF
        i = i + 1
        Me.cboWindow.AddItem rsTemp!����
        If frm��ҩ����.Lvw.SelectedItem.SubItems(5) = rsTemp!���� Then
            intRow = i
        End If
        rsTemp.MoveNext
    Loop
    
    If frm��ҩ����.Lvw.SelectedItem Is Nothing Then Exit Sub
    
    If frm��ҩ����.Lvw.SelectedItem.SubItems(5) <> "" Then
        If intRow >= 0 Then
            cboWindow.ListIndex = intRow
        Else
            Me.cboWindow.AddItem frm��ҩ����.Lvw.SelectedItem.SubItems(5)
            cboWindow.ListIndex = i + 1
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt����_GotFocus()
    GetFocus Txt����
End Sub

Private Sub Txt����_GotFocus()
    GetFocus Txt����
End Sub

Private Sub Txt����_KeyPress(KeyAscii As Integer)
    If InStr(1, "!@#$%^&*(){}[];:,.<>?/|\����������������������������%����&����", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Txtҩ��_GotFocus()
    GetFocus Txtҩ��
End Sub

Private Sub Txtҩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CompareStr  As String, RecOpen As New ADODB.Recordset, StrBit As Byte
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txtҩ��) = "" Then
        Txtҩ��.Tag = 0
        Exit Sub
    End If
    
    CompareStr = UCase(Txtҩ��)
    If Mid(CompareStr, 1, 1) = "��" Then
        If InStr(2, CompareStr, "��") <> 0 Then
            CompareStr = Mid(CompareStr, 2, InStr(2, CompareStr, "��") - 2)
        Else
            CompareStr = Mid(CompareStr, 2)
        End If
    End If

    StrBit = GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0")
    
    gstrSQL = " Select ID,����,���� From ���ű� Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (" & _
      " Select distinct ����ID From ��������˵��" & _
      " Where �������� Like '%ҩ��')" & _
      " And To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01'" & _
      " And (���� like [1] Or ���� like [1] or ���� like [1])"
    Set RecOpen = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡҩ��]", IIf(StrBit = "0", "%", "") & CompareStr & "%")
    
    With RecOpen
        If .EOF Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            Me.Txtҩ�� = ""
            Txtҩ��.Tag = 0
            KeyCode = 0
            Exit Sub
        End If
        If .RecordCount > 1 Then
            With FrmMutilSelect
                Set .gRecCommon = RecOpen.Clone
                .gStrHideCol = "000,1000,1500"
                .strCaption = "ҩ��ѡ����"
                .FrmHeight = 3680
                .FrmWidth = 6000
                .Show 1, Me
                
                If .BlnSelect = False Then
                    Unload FrmMutilSelect
                    KeyCode = 0
                    Me.Txtҩ�� = ""
                    Txtҩ��.Tag = 0
                    Exit Sub
                Else
                    Me.Txtҩ�� = .gRecCommon!����
                    Txtҩ��.Tag = .gRecCommon!Id
                    Unload FrmMutilSelect
                End If
            End With
        Else
            Me.Txtҩ�� = !����
            Txtҩ��.Tag = !Id
        End If
        
    End With
    
    Call LoadWindow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetMaxCode(ByVal lngҩ��ID As Long) As String
    Dim StrCode As String
    Dim RecCode As New ADODB.Recordset
    
    
'        If .State = 1 Then .Close
'        gstrSQL = "Select Max(����) Code From ��ҩ���� Where ҩ��ID=" & lngҩ��ID
'
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        .Open gstrSQL, gcnOracle
'        Call SQLTest
    On Error GoTo errHandle
    gstrSQL = "Select Max(����) Code From ��ҩ���� Where ҩ��ID=[1]"
    Set RecCode = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
    
   With RecCode
        If .EOF Then
            GetMaxCode = 1
        Else
            If IsNull(!Code) Then
                GetMaxCode = 1
            Else
                GetMaxCode = !Code + 1
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
