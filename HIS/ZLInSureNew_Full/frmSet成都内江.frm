VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSet�ɶ��ڽ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab stab 
      Height          =   2865
      Left            =   255
      TabIndex        =   24
      Top             =   765
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&A)"
      TabPicture(0)   =   "frmSet�ɶ��ڽ�.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEdit(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEdit(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl������"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbo������"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtPort"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "ǰ�û�����(&Q)"
      TabPicture(1)   =   "frmSet�ɶ��ڽ�.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraҽ��������"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra 
         Caption         =   "ҽ����������"
         Height          =   1380
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   615
         Width           =   5490
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   4150
            MaxLength       =   2
            TabIndex        =   28
            Top             =   240
            Width           =   360
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   300
            Left            =   5100
            TabIndex        =   6
            Top             =   930
            Width           =   300
         End
         Begin VB.TextBox txtIP 
            Height          =   300
            Left            =   870
            TabIndex        =   3
            Top             =   600
            Width           =   4545
         End
         Begin VB.TextBox Txt�˿ں� 
            Height          =   300
            Left            =   870
            TabIndex        =   1
            Top             =   240
            Width           =   1290
         End
         Begin VB.TextBox txtFile 
            Height          =   300
            Left            =   870
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "C:\"
            Top             =   945
            Width           =   4545
         End
         Begin VB.Label Lbl���� 
            Caption         =   "����������"
            Height          =   255
            Left            =   3050
            TabIndex        =   27
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblIn 
            AutoSize        =   -1  'True
            Caption         =   "&IP��ַ"
            Height          =   180
            Index           =   2
            Left            =   300
            TabIndex        =   2
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblIn 
            AutoSize        =   -1  'True
            Caption         =   "�˿ں�"
            Height          =   180
            Index           =   1
            Left            =   300
            TabIndex        =   0
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblIn 
            AutoSize        =   -1  'True
            Caption         =   "�����ļ�"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   1005
            Width           =   720
         End
      End
      Begin VB.TextBox TxtPort 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4305
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "1"
         Top             =   2220
         Width           =   360
      End
      Begin VB.ComboBox cbo������ 
         Height          =   300
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2220
         Width           =   1770
      End
      Begin VB.Frame fraҽ�������� 
         Caption         =   "ҽԺǰ��ҽ��������"
         Height          =   1545
         Left            =   -74610
         TabIndex        =   25
         Top             =   720
         Width           =   4965
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   13
            Top             =   330
            Width           =   2385
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1260
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   720
            Width           =   2385
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   17
            Top             =   1110
            Width           =   2385
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "����(&T)"
            Height          =   1095
            Left            =   3870
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�û���(&U)"
            Height          =   180
            Index           =   0
            Left            =   390
            TabIndex        =   12
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&P)"
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   14
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "������(&S)"
            Height          =   180
            Index           =   2
            Left            =   390
            TabIndex        =   16
            Top             =   1170
            Width           =   810
         End
      End
      Begin VB.Label lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������(&R)"
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)"
         Height          =   180
         Index           =   3
         Left            =   3300
         TabIndex        =   9
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ŵ���"
         Height          =   180
         Index           =   4
         Left            =   4695
         TabIndex        =   11
         Top             =   2280
         Width           =   540
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   3765
      Width           =   6600
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   15
      TabIndex        =   22
      Top             =   615
      Width           =   6375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3765
      TabIndex        =   20
      Top             =   3915
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5040
      TabIndex        =   19
      Top             =   3915
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   5835
      Top             =   -45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "�������������ҽ��ǰ�÷����������������������̡�"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   21
      Top             =   315
      Width           =   7125
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet�ɶ��ڽ�.frx":0038
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmSet�ɶ��ڽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
End Enum


Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSel_Click()
    Dim strFile As String
    
    Err = 0
    On Error Resume Next
    With dlg
        .Filter = "�����ļ�(*.ini)|*.ini;*.txt"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Sub
        strFile = .FileName
    End With
    Err = 0
    txtFile.Text = strFile
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Text) = False Then
        Exit Sub
    End If
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub


Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    Dim i As Integer
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
     
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�ɶ��ڽ�
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "ҽ���û���"
                  txtEdit(textҽ���û�).Text = Nvl(!����ֵ)
            Case "ҽ���û�����"
                  txtEdit(Textҽ������).Text = Nvl(!����ֵ)
            Case "ҽ��������"
                  txtEdit(Textҽ��������).Text = Nvl(!����ֵ)
            Case "����������"
                  txt����.Text = IIf(IsNull(!����ֵ), 0, !����ֵ)
            End Select
            .MoveNext
        Loop
    End With
    GetRegInFor g����ȫ��, "ҽ��", "���ں�", strReg
    Me.txtPort.Text = IIf(strReg = "", 1, strReg)
    
    GetRegInFor g����ȫ��, "ҽ��", "������", strReg
    For i = 0 To cbo������.ListCount - 1
        If i = Val(strReg) Then
            cbo������.ListIndex = i
        End If
    Next
    
    GetRegInFor g����ȫ��, "ҽ��", "ConfigFileName", strReg
    txtFile.Text = strReg
    GetRegInFor g����ȫ��, "ҽ��", "HostPort", strReg
    txt�˿ں�.Text = strReg
    GetRegInFor g����ȫ��, "ҽ��", "IPAddress", strReg
    txtIP.Text = strReg
    
 End Sub
Private Sub Form_Load()
    mblnFirst = True
    Call LoadBaseData
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If txtFile.Text = "" Then
'        MsgBox "�����ļ�δѡ��", vbInformation + vbDefaultButton1, gstrSysName
'        stab.Tab = 0
'        If txtFile.Enabled Then txtFile.SetFocus
'        Exit Function
    End If
    
    If Dir(txtFile.Text) <> "" Then
    Else
        MsgBox "�����ļ�������", vbInformation + vbDefaultButton1, gstrSysName
        stab.Tab = 0
        If txtFile.Enabled Then txtFile.SetFocus
        Exit Function
    End If
    If Trim(txtIP.Text) = "" Then
        MsgBox "IPδ����", vbInformation + vbDefaultButton1, gstrSysName
        stab.Tab = 0
        If txtIP.Enabled Then txtIP.SetFocus
        Exit Function
    End If
    If Trim(txt�˿ں�.Text) = "" Then
        MsgBox "�˿ں�δ����", vbInformation + vbDefaultButton1, gstrSysName
        stab.Tab = 0
        If txt�˿ں�.Enabled Then txt�˿ں�.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    Dim lngRow As Long
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�ɶ��ڽ� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ��ڽ� & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ��ڽ� & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ��ڽ� & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�ɶ��ڽ� & ",null,'����������','" & txt����.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    SaveRegInFor g����ȫ��, "ҽ��", "������", Split(cbo������.Text, "-")(0)
    SaveRegInFor g����ȫ��, "ҽ��", "���ں�", Val(txtPort.Text)
    SaveRegInFor g����ȫ��, "ҽ��", "ConfigFileName", txtFile.Text
    SaveRegInFor g����ȫ��, "ҽ��", "HostPort", txt�˿ں�.Text
    SaveRegInFor g����ȫ��, "ҽ��", "IPAddress", txtIP.Text
    
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
End Sub
Private Function Open�м��() As Boolean
    '�����м��
    '�м������
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String
    
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�山ҽ��", TYPE_�ɶ��ڽ�)
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_�ɶ��ڽ� = New ADODB.Connection

    If OraDataOpen(gcnOracle_�ɶ��ڽ�, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    Open�м�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub LoadBaseData()
    '��������
    With cbo������
        .Clear
        .AddItem "0-��������������"
        .ListIndex = .NewIndex
        .AddItem "1-��ɭ����������"
    End With
End Sub



Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub

Private Sub txtIP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub TxtPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stab.Tab = 1
        If txtEdit(textҽ���û�).Enabled Then txtEdit(textҽ���û�).SetFocus
    End If
End Sub

Private Sub TxtPort_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtPort, KeyAscii, m����ʽ
End Sub

Public Function ��������() As Boolean
'���ܣ������붫�󰢶��ɵ�ҽ���ӿ�

    
    mblnOK = False
    
    On Error GoTo errHandle
    mblnChange = False
    frmSet�ɶ��ڽ�.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Txt�˿ں�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
        zlControl.TxtCheckKeyPress txt����, KeyAscii, m����ʽ
End Sub
