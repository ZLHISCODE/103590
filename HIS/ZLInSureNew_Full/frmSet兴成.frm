VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSet�˳� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTab1 
      Height          =   3270
      Left            =   120
      TabIndex        =   4
      Top             =   915
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   5768
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "ҽԺ��ǰ�û�(&0)"
      TabPicture(0)   =   "frmSet�˳�.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraҽ��������(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ҽ����ǰ�û�(&1)"
      TabPicture(1)   =   "frmSet�˳�.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdZXTest"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdODBC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraҽ��������(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "����(&2)"
      TabPicture(2)   =   "frmSet�˳�.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lbl(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lbl(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lbl(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lbl(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lbl(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lbl(5)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtPath(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtPath(1)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtPath(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtPath(3)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdSel(0)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdSel(1)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdSel(3)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmdSel(2)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cmdSel(4)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtPath(4)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "chk������"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cboҽԺ����"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Chk�����������"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
      Begin VB.CheckBox Chk����������� 
         Caption         =   "�����������(&Q)"
         Height          =   285
         Left            =   3795
         TabIndex        =   41
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.ComboBox cboҽԺ���� 
         Height          =   300
         Left            =   1320
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   2610
         Width           =   2340
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "��վ����ڶ�����(&R)"
         Height          =   285
         Left            =   3795
         TabIndex        =   38
         Top             =   2580
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CommandButton cmdZXTest 
         Caption         =   "����(&T)"
         Height          =   350
         Left            =   -70485
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1100
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   4
         Left            =   1305
         TabIndex        =   28
         Text            =   "C:\"
         Top             =   2235
         Width           =   4020
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   300
         Index           =   4
         Left            =   5370
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2250
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   300
         Index           =   2
         Left            =   5370
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1440
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   300
         Index           =   3
         Left            =   5370
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1860
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   300
         Index           =   1
         Left            =   5370
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1020
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   300
         Index           =   0
         Left            =   5370
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   285
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   3
         Left            =   1305
         TabIndex        =   22
         Text            =   "C:\Out"
         Top             =   1845
         Width           =   4020
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   2
         Left            =   1305
         TabIndex        =   20
         Text            =   "C:\IN"
         Top             =   1440
         Width           =   4020
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   1
         Left            =   1305
         TabIndex        =   18
         Text            =   "C:\xcyb\put"
         Top             =   1035
         Width           =   4020
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Index           =   0
         Left            =   1305
         TabIndex        =   16
         Text            =   "C:\xcyb\get"
         Top             =   600
         Width           =   4020
      End
      Begin VB.CommandButton cmdODBC 
         Caption         =   "����Դ(&D)"
         Height          =   350
         Left            =   -70485
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1100
      End
      Begin VB.Frame fraҽ�������� 
         Height          =   1875
         Index           =   0
         Left            =   -74820
         TabIndex        =   5
         Top             =   660
         Width           =   5595
         Begin VB.CommandButton cmdTest 
            Caption         =   "����(&T)"
            Height          =   1095
            Left            =   4515
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   555
            Width           =   1005
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   8
            Top             =   1335
            Width           =   3075
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1200
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   945
            Width           =   3075
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   6
            Top             =   555
            Width           =   3075
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "������(&S)"
            Height          =   180
            Index           =   2
            Left            =   330
            TabIndex        =   12
            Top             =   1395
            Width           =   810
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "����(&P)"
            Height          =   180
            Index           =   1
            Left            =   510
            TabIndex        =   11
            Top             =   1005
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�û���(&U)"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   10
            Top             =   615
            Width           =   810
         End
      End
      Begin VB.Frame fraҽ�������� 
         Height          =   2070
         Index           =   1
         Left            =   -74850
         TabIndex        =   30
         Top             =   705
         Width           =   5595
         Begin VB.TextBox txtODBC 
            Height          =   300
            Index           =   0
            Left            =   1710
            MaxLength       =   40
            TabIndex        =   33
            Top             =   555
            Width           =   2565
         End
         Begin VB.TextBox txtODBC 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1710
            MaxLength       =   40
            TabIndex        =   32
            Top             =   930
            Width           =   2565
         End
         Begin VB.TextBox txtODBC 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1710
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   1275
            Width           =   2565
         End
         Begin VB.Label lblODBC 
            AutoSize        =   -1  'True
            Caption         =   "ODBC����Դ��(&U)"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   36
            Top             =   615
            Width           =   1350
         End
         Begin VB.Label lblODBC 
            AutoSize        =   -1  'True
            Caption         =   "ODBC����Դ�û�(&U)"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   35
            Top             =   975
            Width           =   1530
         End
         Begin VB.Label lblODBC 
            AutoSize        =   -1  'True
            Caption         =   "ODBC����Դ����(&P)"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   34
            Top             =   1380
            Width           =   1530
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽԺ����"
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   39
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����ϵͳĿ¼"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   29
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ӿڳ���Ŀ¼"
         Height          =   180
         Index           =   3
         Left            =   225
         TabIndex        =   21
         Top             =   1890
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ӿ����Ŀ¼"
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   19
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label lbl 
         Caption         =   "�ϴ���ʱĿ¼"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   1110
         Width           =   1140
      End
      Begin VB.Label lbl 
         Caption         =   "������ʱĿ¼"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   15
         Top             =   675
         Width           =   1140
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   3
      Top             =   4290
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   7665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4065
      TabIndex        =   0
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5250
      TabIndex        =   1
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Label lblNote 
      Caption         =   "    ���õ�ҽ�Ʊ������ݷ����������Ӵ���Ϊ��֤������Ч����ʱҽ�Ʊ������ݷ�����������á�"
      Height          =   390
      Left            =   810
      TabIndex        =   14
      Top             =   240
      Width           =   5475
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet�˳�.frx":0054
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmSet�˳�"
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
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private mҽԺ���� As String

Public Function ��������() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSet�˳�.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdODBC_Click()
    On Error Resume Next
    Shell "ODBCAD32", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "���ܽ���ODBC����Դ������������ϵͳ�Ƿ���ȷ��װ��", vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Dim strPath As String
    strPath = OpenDire(Me, "��ָ��Ŀ¼��")
    If strPath = "" Then Exit Sub
    txtPath(Index).Text = strPath
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub cmdZXTest_Click()
   Dim cnInsure As New ADODB.Connection
    Err = 0
    On Error Resume Next
    With cnInsure
            If .State = adStateOpen Then .Close
            .ConnectionString = "dsn=" & txtODBC(0).Text & ";uid=" & txtODBC(1).Text & ";pwd=" & txtODBC(2).Text & ""
            .Open
            If Err <> 0 Then
                MsgBox "���Բ��ɹ�������ҽ�����ݷ������Ƿ���ã��Լ�����Դ�Ƿ���ȷ���ã�", vbExclamation, gstrSysName
                Exit Sub
            End If
            .Close
            MsgBox "���Գɹ�����ҽ�����ݷ������������ӣ�", vbInformation, gstrSysName
    End With
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    
    mblnFirst = False
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�˳ɺ˹�ҵ
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    mҽԺ���� = "01"
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "ҽ���û���"
                  txtEdit(textҽ���û�).Text = Nvl(!����ֵ)
            Case "ҽ���û�����"
                  txtEdit(Textҽ������).Text = Nvl(!����ֵ)
            Case "ҽ��������"
                  txtEdit(Textҽ��������).Text = Nvl(!����ֵ)
            Case "ҽԺ����"
                  mҽԺ���� = Nvl(!����ֵ, "01")
            End Select
            .MoveNext
        Loop
    End With
    txtODBC(0).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_NAME"), "")
    txtODBC(1).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_USERNAME"), "")
    txtODBC(2).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_PASSWORD"), "")
    
    txtPath(0).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Get"), "C:\xcyb\get")
    txtPath(1).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Put"), "C:\xcyb\Put")
    txtPath(2).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_In"), "C:\xcyb\In")
    txtPath(3).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Out"), "C:\xcyb\Out")
    txtPath(4).Text = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_System"), "C:\")
    
    chk������.Value = IIf(Val(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("������"), "1")) = 1, 1, 0)
    
    '�º�����20050408����
    
    Chk�����������.Value = IIf(Val(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("�����������"), "1")) = 1, 1, 0)
    
    Call LoadҽԺ����
 
 
 End Sub

Private Sub Form_Load()
    mblnFirst = True
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
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�˳ɺ˹�ҵ & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˳ɺ˹�ҵ & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˳ɺ˹�ҵ & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˳ɺ˹�ҵ & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˳ɺ˹�ҵ & ",null,'ҽԺ����','" & Split(cboҽԺ����.Text, " ")(0) & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Get"), Trim(txtPath(0).Text)
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Put"), Trim(txtPath(1).Text)
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_In"), Trim(txtPath(2).Text)
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Out"), Trim(txtPath(3).Text)
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_System"), Trim(txtPath(4).Text)
    
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_NAME"), Trim(txtODBC(0).Text)
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_USERNAME"), Trim(txtODBC(1).Text)
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_PASSWORD"), Trim(txtODBC(2).Text)
    
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("������"), IIf(chk������.Value = 1, 1, 0)
    
    '�º�����20050408�������޸�,����סԺ����ʱ����ҽ������,���Ǹò����ڶ�̬���ʼ����һ������˿�
    
    SaveSetting "ZLSOFT", "����ģ��\zl9Insure", UCase("�����������"), IIf(Chk�����������.Value = 1, 1, 0)

    
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
Private Function OpenDire(odtvOwner As Form, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDire = sBuffer
   End If
End Function

Private Function CreatePath(ByVal strPath As String) As Boolean
    '����:�����ļ�·��
    Dim objPath As New FileSystemObject
    Dim strArr As Variant
    Dim strTmpPath As String
    CreatePath = False
    Dim i As Long
    strTmpPath = strPath
    If InStr(strTmpPath, "\") = 0 Or InStr(strTmpPath, "\\") <> 0 Then
         MsgBox "·������ȷ!", vbInformation + vbDefaultButton1, gstrSysName
         Exit Function
    End If
    strArr = Split(strTmpPath, "\")
    If InStr(1, "A:B:C:D:E:F:G:H:I:J:K:L:M:N:O:P:Q:R:S:T:U:V:W:X:Y:Z:", strArr(0)) = 0 Then
         MsgBox "·������ȷ!", vbInformation + vbDefaultButton1, gstrSysName
         Exit Function
    End If
    
    strTmpPath = strArr(0)
    For i = 1 To UBound(strArr)
        Err = 0
        On Error Resume Next
        strTmpPath = strTmpPath & "\" & strArr(i)
        
        If objPath.FolderExists(strTmpPath) = False Then
            objPath.CreateFolder strTmpPath
            If Err <> 0 Then
                MsgBox "����·��ʧ��(" & strTmpPath & ")" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ��������:" & Err.Description, vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        End If
    Next
    CreatePath = True
End Function


Private Sub txtODBC_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtPath_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtPath_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(1, ";-`@#$%^&**()#@+-|", Chr(KeyAscii)) <> 0 Then
        KeyAscii = 0
    End If
End Sub
Private Function LoadҽԺ����()
    Dim i As Integer
    cboҽԺ����.Clear
    
    With cboҽԺ����
        .AddItem "01 һ��ҽԺ"
        .AddItem "02 ����ҽԺ"
        .AddItem "03 ����ҽԺ"
        .AddItem "11 ת�������һ��ҽԺ"
        .AddItem "12 ת���ʡ��һ��ҽԺ"
        .AddItem "13 ת���ʡ��һ��ҽԺ"
        .AddItem "14 ת������ڶ���ҽԺ"
        .AddItem "15 ת���ʡ�ڶ���ҽԺ"
        .AddItem "16 ת���ʡ�����ҽԺ"
        .AddItem "17 ת�����������ҽԺ"
        .AddItem "18 ת���ʡ������ҽԺ"
        .AddItem "19 ת���ʡ������ҽԺ"
        
        For i = 0 To .ListCount - 1
            If Split(.List(i), " ")(0) = mҽԺ���� Then
                .ListIndex = i
                Exit For
            End If
        Next
        If .ListIndex < 0 Then
            .ListIndex = 0
        End If
    End With
End Function

