VERSION 5.00
Begin VB.Form frmTendItemTemplateEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀģ��༭"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   Icon            =   "frmTendItemTemplateEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5835
      Left            =   5340
      TabIndex        =   16
      Top             =   -270
      Width           =   45
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5550
      TabIndex        =   15
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5550
      TabIndex        =   14
      Top             =   300
      Width           =   1100
   End
   Begin VB.ComboBox cbo���û���ȼ� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   570
      Width           =   2655
   End
   Begin VB.TextBox txtģ������ 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   180
      Width           =   2655
   End
   Begin VB.PictureBox picCloumn 
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   60
      ScaleHeight     =   3405
      ScaleWidth      =   5205
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1260
      Width           =   5205
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2130
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2310
         Width           =   975
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "����(&U)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   2130
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2010
         Width           =   975
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   8
         Top             =   465
         Width           =   1770
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2130
         TabIndex        =   12
         Top             =   885
         Width           =   975
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2130
         TabIndex        =   13
         Top             =   1185
         Width           =   975
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   2760
         Left            =   3240
         TabIndex        =   9
         Top             =   450
         Width           =   1770
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ�����¼��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   630
      TabIndex        =   4
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lbl���û���ȼ� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ȼ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   630
      Width           =   720
   End
   Begin VB.Label lblģ������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ģ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmTendItemTemplateEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mblnSel As String
Private mblnStart As Boolean
Private mlng����ID As Long
Private mint����ȼ� As Integer
Private mstrģ������ As String
Private mblnEdit As Boolean

Public Function ShowEditor(ByVal objParent As Object, ByVal lng����id As Long, ByVal strģ������ As String, ByVal int����ȼ� As Integer) As String
    On Error Resume Next
    mblnSel = ""
    mblnEdit = False
    mlng����ID = lng����id
    mstrģ������ = strģ������
    mint����ȼ� = int����ȼ�
    Me.Show 1, objParent
    ShowEditor = mblnSel
End Function

Private Sub cbo����_Click()
    Call cbo���û���ȼ�_Click
End Sub

Private Sub cbo���û���ȼ�_Click()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If mblnStart = False Then Exit Sub
    
    gstrSQL = " Select A.��Ŀ���,A.��Ŀ���� From �����¼��Ŀ A" & _
              " Where A.Ӧ�÷�ʽ<>0 " & IIf(cbo���û���ȼ�.ItemData(cbo���û���ȼ�.ListIndex) = -1, "", " And A.����ȼ�>=[2]") & _
              " And (A.���ÿ���=1 Or (A.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=A.��Ŀ��� And D.����id=[1])))" & _
              " MINUS " & _
              " Select B.��Ŀ���,B.��Ŀ���� From ������Ŀģ�� A,�����¼��Ŀ B " & _
              " Where A.��Ŀ���=B.��Ŀ��� And A.����ID =[1] And A.����ȼ�=[2]" & _
              " Order by ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cbo����.ItemData(Me.cbo����.ListIndex), Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex))
    
    With rsTemp
        Me.lstColumnItems.Clear
        Do While Not .EOF
            Me.lstColumnItems.AddItem !��Ŀ����
            Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = !��Ŀ���
            .MoveNext
        Loop
    End With
    
    '��ȡ��ѡ�����Ŀ�嵥
    gstrSQL = " Select B.��Ŀ���,B.��Ŀ���� From ������Ŀģ�� A,�����¼��Ŀ B " & _
              " Where A.��Ŀ���=B.��Ŀ��� And A.����ID =[1] And A.����ȼ�=[2]" & _
              " Order by A.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cbo����.ItemData(Me.cbo����.ListIndex), Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex))
    
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.lstColumnUsed.AddItem !��Ŀ����
            Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = !��Ŀ���
            .MoveNext
        Loop
    End With

    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0)
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim intIndex As Integer
    Dim objlst As ListBox
    If Index = 0 Then
        If Me.lstColumnItems.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnItems.ListIndex
        Me.lstColumnUsed.AddItem Me.lstColumnItems.Text
        Me.lstColumnUsed.ItemData(Me.lstColumnUsed.NewIndex) = Me.lstColumnItems.ItemData(Me.lstColumnItems.ListIndex)
        Me.lstColumnItems.RemoveItem Me.lstColumnItems.ListIndex
        Set objlst = lstColumnItems
    Else
        If Me.lstColumnUsed.ListIndex < 0 Then Exit Sub
        intIndex = Me.lstColumnUsed.ListIndex
        Me.lstColumnItems.AddItem Me.lstColumnUsed.Text
        Me.lstColumnItems.ItemData(Me.lstColumnItems.NewIndex) = Me.lstColumnUsed.ItemData(Me.lstColumnUsed.ListIndex)
        Me.lstColumnUsed.RemoveItem Me.lstColumnUsed.ListIndex
        Set objlst = lstColumnUsed
    End If
    If objlst.ListCount >= intIndex + 1 Then
        objlst.ListIndex = intIndex
    Else
        objlst.ListIndex = objlst.ListCount - 1
    End If
    
    cmdColumn(0).Enabled = (lstColumnItems.ListCount <> 0)
    cmdColumn(1).Enabled = (lstColumnUsed.ListCount <> 0)
    mblnEdit = True
    
    Call SetMoveState
End Sub

Private Sub cmdMove_Click(Index As Integer)
    Dim arrData
    Dim strCopy As String
    Dim lngDo As Long, lngMax As Long
    Dim lngSelIndex As Long, lngTarIndex As Long
    
    '��ǰ����
    lngSelIndex = lstColumnUsed.ListIndex
    'Ŀ������
    lngTarIndex = lngSelIndex + IIf(Index = 0, -1, 1)
    lngMax = lstColumnUsed.ListCount - 1
    For lngDo = 0 To lngMax
        If lngDo = lngTarIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngSelIndex) & "," & lstColumnUsed.ItemData(lngSelIndex)
        ElseIf lngDo = lngSelIndex Then
            strCopy = strCopy & "|" & lstColumnUsed.List(lngTarIndex) & "," & lstColumnUsed.ItemData(lngTarIndex)
        Else
            strCopy = strCopy & "|" & lstColumnUsed.List(lngDo) & "," & lstColumnUsed.ItemData(lngDo)
        End If
    Next
    strCopy = Mid(strCopy, 2)
    Debug.Print strCopy
    
    lstColumnUsed.Clear
    arrData = Split(strCopy, "|")
    For lngDo = 0 To lngMax
        lstColumnUsed.AddItem Split(arrData(lngDo), ",")(0)
        lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Split(arrData(lngDo), ",")(1))
    Next
    lstColumnUsed.ListIndex = lngTarIndex
    Call SetMoveState
End Sub

Private Sub cmdOK_Click()
    Dim blnTrans As Boolean
    Dim intRow As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Me.lstColumnUsed.ListCount = 0 Then
        MsgBox "��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txtģ������.Text) = "" Then
        MsgBox "��¼��ģ�����ƣ�", vbInformation, gstrSysName
        txtģ������.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txtģ������.Text, vbFromUnicode)) > 50 Then
        MsgBox "ģ�����Ƴ������������25�����ֻ�50���ַ���", vbInformation, gstrSysName
        txtģ������.SetFocus
        Exit Sub
    End If
    
    '��������������Ƿ���ڸ�ģ��
    If mint����ȼ� = 9 Then
        gstrSQL = " Select 1 From ������Ŀģ�� Where ����ȼ�=[1] And ����ID=[2] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex), Me.cbo����.ItemData(Me.cbo����.ListIndex))
        If rsTemp.RecordCount <> 0 Then
            If MsgBox("�Ѵ��ڸû�����Ŀģ��,�㡰�ǡ�����£�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    '׼������
    intCount = Me.lstColumnUsed.ListCount
    gcnOracle.BeginTrans
    blnTrans = True
    
    Call zlDatabase.ExecuteProcedure("zl_������Ŀģ��_Delete(" & Me.cbo����.ItemData(Me.cbo����.ListIndex) & "," & Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex) & ")", "ɾ����ǰģ��")
    For intRow = 1 To intCount
        Debug.Print "zl_������Ŀģ��_Insert(" & cbo����.ItemData(Me.cbo����.ListIndex) & ",'" & Me.txtģ������.Text & "'," & Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex) & "," & Me.lstColumnUsed.ItemData(intRow - 1) & "," & intRow & ")"
        Call zlDatabase.ExecuteProcedure("zl_������Ŀģ��_Insert(" & cbo����.ItemData(Me.cbo����.ListIndex) & ",'" & Me.txtģ������.Text & "'," & Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex) & "," & Me.lstColumnUsed.ItemData(intRow - 1) & "," & intRow & ")", "����ģ������")
    Next
    
    gcnOracle.CommitTrans
    blnTrans = False
    mblnSel = Me.cbo���û���ȼ�.ItemData(Me.cbo���û���ȼ�.ListIndex) & "|" & Me.cbo����.ItemData(Me.cbo����.ListIndex)
    
    mblnEdit = False
    Unload Me
    Exit Sub
errHand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        blnTrans = False
    End If
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        mblnEdit = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'װ��ȱʡ����
    mblnStart = False
    
    With Me.cbo���û���ȼ�
        .Clear
        .AddItem "�ؼ�����¼��ģ��"
        .ItemData(.NewIndex) = 0
        .AddItem "һ������¼��ģ��"
        .ItemData(.NewIndex) = 1
        .AddItem "��������¼��ģ��"
        .ItemData(.NewIndex) = 2
        .AddItem "��������¼��ģ��"
        .ItemData(.NewIndex) = 3
        .AddItem "����/����¼��ģ��"
        .ItemData(.NewIndex) = -1
        .ListIndex = 0
    End With
    
    '��ȡ�ٴ�����(���ǵ����ܻ�ʿ���ڲ���,����ҳ���Ӧ�Ŀ���,������û�ʿ����;�ٿ��ǵ��ٴ����ҵ���ֱ��������ģ��,����ģʽ��֧��)
    If InStr(1, mstrPrivs, "�༭��������ģ��") <> 0 Then
        gstrSQL = " Select B.ID,B.���� " & _
                  " From ��������˵�� A,���ű� B" & _
                  " Where A.��������='�ٴ�' And A.������� IN (2,3) And A.����ID=B.ID" & _
                  " Order by B.����"
    Else
        gstrSQL = " Select B.ID,B.����,B.���� " & _
                  " From ��������˵�� A,���ű� B,������Ա C" & _
                  " Where A.��������='�ٴ�' And A.������� IN (2,3) And A.����ID=B.ID" & _
                  " And B.ID=C.����ID And C.��ԱID=[1]" & _
                  " UNION " & _
                  " Select B.ID,B.����,B.���� " & _
                  " From ��������˵�� A,���ű� B,�������Ҷ�Ӧ C" & _
                  " Where A.��������='�ٴ�' And A.������� IN (2,3) And A.����ID=B.ID And B.ID=C.����ID And C.����ID=[2]"
        gstrSQL = " Select Distinct ID,����,���� From (" & gstrSQL & ") Order by ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId, glngDeptId)
    With rsTemp
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem !����
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = !ID
            If !ID = mlng����ID Then Me.cbo����.ListIndex = .AbsolutePosition - 1
            .MoveNext
        Loop
        If .RecordCount = 0 Then
            MsgBox "�㲻�����κ�һ���ٴ����ң�", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
        
        If Me.cbo����.ListIndex = -1 Then Me.cbo����.ListIndex = 0
    End With
    mblnStart = True
    
    Me.txtģ������.Text = mstrģ������
    If mint����ȼ� = 9 Then
        Me.cbo���û���ȼ�.Enabled = True
    Else
        If mint����ȼ� = -1 Then
            Me.cbo���û���ȼ�.ListIndex = 4
        Else
            Me.cbo���û���ȼ�.ListIndex = mint����ȼ�
        End If
        Me.cbo���û���ȼ�.Enabled = False
        Me.cbo����.Enabled = False
    End If
    Call cbo���û���ȼ�_Click
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("��δ�������ݣ��Ƿ��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub lstColumnItems_DblClick()
    If lstColumnItems.ListCount = 0 Then Exit Sub
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    Call SetMoveState
End Sub

Private Sub lstColumnUsed_DblClick()
    If lstColumnUsed.ListCount = 0 Then Exit Sub
    Call cmdColumn_Click(1)
End Sub

Private Sub txtģ������_GotFocus()
    Call zlControl.TxtSelAll(txtģ������)
End Sub

Private Sub SetMoveState()
    cmdMove(0).Enabled = False
    cmdMove(1).Enabled = False
    
    If lstColumnUsed.ListIndex < 0 Then Exit Sub
    If lstColumnUsed.SelCount < 0 Then Exit Sub
    cmdMove(0).Enabled = (lstColumnUsed.ListIndex > 0)
    cmdMove(1).Enabled = (lstColumnUsed.ListIndex < lstColumnUsed.ListCount - 1)
End Sub
