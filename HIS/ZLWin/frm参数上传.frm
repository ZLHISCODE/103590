VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm�����ϴ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ϴ�"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frm�����ϴ�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ListView lvw������ 
      Height          =   2610
      Left            =   1110
      TabIndex        =   30
      Top             =   3735
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilt16"
      SmallIcons      =   "ilt16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "������"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��������"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�ϴ��û�"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�ϴ�վ��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "��������"
         Object.Width           =   4304
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   5055
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5070
      TabIndex        =   26
      Top             =   195
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   27
      Top             =   600
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ilt32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ϴ�.frx":000C
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ϴ�.frx":0326
            Key             =   "Client"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ϴ�.frx":0DF0
            Key             =   "Scheame"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilt16 
      Left            =   2900
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ϴ�.frx":110A
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ϴ�.frx":1424
            Key             =   "Client"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ϴ�.frx":1EEE
            Key             =   "Scheame"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "�ϴ�����"
      Height          =   3570
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4830
      Begin VB.ComboBox cbo��ʽ 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2745
         Width           =   3600
      End
      Begin VB.ComboBox cbo�û� 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3120
         Width           =   3600
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   930
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "������"
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   930
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "��������"
         Top             =   720
         Width           =   3570
      End
      Begin VB.TextBox txtEdit 
         Height          =   1590
         Index           =   3
         Left            =   930
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Tag             =   "��������"
         Top             =   1095
         Width           =   3570
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   285
         Left            =   4470
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�ϴ���ʽ"
         Height          =   180
         Index           =   3
         Left            =   165
         TabIndex        =   8
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����û�"
         Height          =   180
         Index           =   4
         Left            =   165
         TabIndex        =   10
         Top             =   3180
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   345
         TabIndex        =   1
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   6
         Top             =   1095
         Width           =   720
      End
   End
   Begin VB.Frame fra 
      Caption         =   "ע����Ϣ������ָ�"
      Height          =   3555
      Index           =   2
      Left            =   90
      TabIndex        =   29
      Top             =   135
      Width           =   4815
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��"
         Height          =   330
         Left            =   4335
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1785
         Width           =   300
      End
      Begin VB.CommandButton cmdBakup 
         Caption         =   "����(&B)"
         Height          =   350
         Left            =   2475
         TabIndex        =   17
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "�ָ�(&R)"
         Height          =   350
         Left            =   3660
         TabIndex        =   18
         Top             =   2400
         Width           =   1100
      End
      Begin VB.TextBox txtFile 
         Height          =   350
         Left            =   825
         MaxLength       =   500
         TabIndex        =   15
         Top             =   1770
         Width           =   3840
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "ע���ļ�"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label lblReg 
         Caption         =   "    ����������Ҫ��ָ��ע�����ZLSOFT�µ����в������ñ��ݳ�ָ����reg�ļ�,�Ա��ָ���"
         Height          =   450
         Index           =   0
         Left            =   915
         TabIndex        =   12
         Top             =   390
         Width           =   3795
      End
      Begin VB.Image img 
         Height          =   480
         Left            =   165
         Picture         =   "frm�����ϴ�.frx":2208
         Top             =   615
         Width           =   480
      End
      Begin VB.Label lblReg 
         Caption         =   "    �����ָ���Ҫ��ָ�����ݵ�Reg�ļ�����ע�����Ϣ�Ļָ���"
         Height          =   435
         Index           =   1
         Left            =   885
         TabIndex        =   13
         Top             =   855
         Width           =   3795
      End
   End
   Begin VB.Frame fra 
      Caption         =   "�����ָ�"
      Height          =   3555
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   150
      Width           =   4830
      Begin VB.Frame Frame3 
         Caption         =   "�û�ѡ��"
         Height          =   2220
         Left            =   105
         TabIndex        =   28
         Top             =   690
         Width           =   4575
         Begin VB.CheckBox chkAllUser 
            Caption         =   "�����û�"
            Height          =   240
            Left            =   3390
            TabIndex        =   22
            Top             =   0
            Width           =   1035
         End
         Begin MSComctlLib.ListView lvw�û� 
            Height          =   1935
            Left            =   90
            TabIndex        =   23
            Top             =   225
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ilt32"
            SmallIcons      =   "ilt16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "�û���"
               Object.Width           =   4304
            EndProperty
         End
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   3720
      End
      Begin VB.Label lblInforLst 
         Caption         =   "������Ϣ���ϴ�վ��[lxh],�ϴ��û�(zlhis)"
         ForeColor       =   &H80000001&
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   3270
         Width           =   4680
      End
      Begin VB.Label lblInfor 
         Caption         =   "�ָ�������[123456789]lxh�����·���"
         ForeColor       =   &H80000001&
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   3030
         Width           =   4680
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ָ���ʽ"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm�����ϴ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrPrvs As String      'Ȩ�޴�
Dim mblnFirst As Boolean
Dim mbytParaType As Byte    '0-�ϴ�,1-����,2-����ע�������ָ�

Public Sub ShowEdit(ByVal frmMain As Object, ByVal bytParaType As Byte)
    '----------------------------------------------------------------------------------------------------------------
    '--����:��ʾ�ϴ�����.
    '--����:frmMain-������
    '       bytParaType-��������(0-�ϴ�,1-����)
    '       strPrivs-Ȩ�޴�
    '----------------------------------------------------------------------------------------------------------------
    mstrPrvs = GetPrivFunc(0, �����嵥.���ز�������)
    mbytParaType = bytParaType
    
    Me.Show vbModal, frmMain
End Sub
Private Sub setCtlShowMode()
    '����:���ÿؼ�����ʾģʽ
    fra(0).Visible = False
    fra(1).Visible = False
    fra(2).Visible = False
    If mbytParaType = 0 Then
        fra(0).Visible = True
    ElseIf mbytParaType = 1 Then
        fra(1).Visible = True
        cmdSave.Enabled = True
    Else
        fra(2).Visible = True
        cmdSave.Enabled = True
    End If
End Sub
Private Sub cbo��ʽ_Click()
    Dim bytType As Byte, i As Integer
    If mbytParaType <> 0 Then Exit Sub
    
    '��ʼ������
    bytType = Val(Split(Me.cbo��ʽ.Text, "-")(0))
    Select Case bytType
    Case 1          '����
        Me.cbo�û�.Enabled = False      '����ѡ�û�
    Case Else         '����,˽��
        'Ҫָ���û�
        If InStr(1, mstrPrvs, "�����ϴ�") <> 0 Then
            Me.cbo�û�.Enabled = True
        Else
            'ָ���ǵ�ǰ�û�
            Me.cbo�û�.ListIndex = -1
            For i = 0 To Me.cbo�û�.ListCount - 1
                If Me.cbo�û�.List(i) = UCase(gstrDbUser) Then
                    Me.cbo�û�.ListIndex = i
                    Exit For
                End If
            Next
            Me.cbo�û�.Enabled = False
        End If
    End Select
End Sub

Private Sub cbo��ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If

End Sub

Private Sub cbo����_Click()
  Dim bytType As Byte, i As Integer
  If mbytParaType = 0 Then Exit Sub
    '��ʼ������
    bytType = Val(Split(Me.cbo����.Text, "-")(0))
    Select Case bytType
    Case 1          '����
        Me.lvw�û�.Enabled = False      '����ѡ�û�
        Me.lvw�û�.BackColor = Me.BackColor
        Me.chkAllUser.Enabled = False
        
    Case Else         '����,˽��
        'Ҫָ���û�
        If InStr(1, mstrPrvs, "��������") <> 0 Then
            Me.lvw�û�.Enabled = True
            Me.lvw�û�.BackColor = Me.cbo����.BackColor
            Me.chkAllUser.Enabled = True
        Else
            'ָ���ǵ�ǰ�û�
            lvw�û�.Enabled = False
            Me.lvw�û�.BackColor = Me.BackColor
            Me.chkAllUser.Enabled = False
        End If
    End Select
End Sub
Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo�û�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkAllUser_Click()
    Dim lstItem As ListItem
    Dim blnCheck As Boolean
    If chkAllUser.Value = 2 Then Exit Sub
    blnCheck = chkAllUser.Value = 1
    For Each lstItem In lvw�û�.ListItems
        lstItem.Checked = blnCheck
    Next
End Sub
Private Sub chkAllUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdBakup_Click()
    '����
    Dim strFile As String
    Dim strPath As String
    Dim strPathAndFile As String
    Dim objFile As New FileSystemObject
    
    If Trim(txtFile.Text) = "" Then
        MsgBox "��ѡ�������Ҫ���ݵ��ļ�!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    strPathAndFile = Trim(txtFile.Text)
    strFile = objFile.GetFileName(strPathAndFile)
    
    strPath = objFile.GetParentFolderName(strPathAndFile)
    If strFile = "" Then
        MsgBox "�������ļ���,������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If FindFile(strPath) = False Then
        MsgBox "�����ڸ�·��,�����������ļ�·��!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    '����
    If ExportReg(strPathAndFile, """HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT""") = False Then
        MsgBox "����ʧ��!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    MsgBox "�����ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRestore_Click()
    '����
    Dim strFile As String
    Dim strPath As String
    Dim strPathAndFile As String
    Dim objFile As New FileSystemObject
    
    If Trim(txtFile.Text) = "" Then
        MsgBox "��ѡ�������Ҫ�ָ����ļ�!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    strPathAndFile = Trim(txtFile.Text)
    strFile = objFile.GetFileName(strPathAndFile)
    strPath = objFile.GetParentFolderName(strPathAndFile)
    If strFile = "" Then
        MsgBox "�����ļ���,������!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If FindFile(strPath) = False Then
        MsgBox "�����ڸ�·��,�����������ļ�·��!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If FindFile(strPathAndFile) = False Then
        MsgBox "�����ڸ�ע���ļ�,�����������ļ�!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    '����
    If RestoreWndowsReg(strPathAndFile) = False Then
        MsgBox "����ʧ��!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    MsgBox "����ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
'    Unload Me
End Sub

Private Sub cmdSave_Click()
   
    If mbytParaType = 0 Then
        '�ϴ�
        If SaveUPData = False Then Exit Sub
        Unload Me
        Exit Sub
    ElseIf mbytParaType = 1 Then
        '����
        If DataBaseToRestoreReg = False Then Exit Sub
        Unload Me
        Exit Sub
    Else
        Unload Me
        Exit Sub
    End If

End Sub
Private Function DataBaseToRestoreReg() As Boolean
    '����:�����õĲ����лָ�������ע�����
        Dim bytType As Byte, lng������ As Long
        Dim cllUser As New Collection, lstItem As ListItem
        DataBaseToRestoreReg = False
        Err = 0: On Error GoTo ErrHand:
        
        bytType = Val(Split(cbo����.Text, "-")(0))
        For Each lstItem In lvw�û�.ListItems
            If lstItem.Checked Then
                cllUser.Add lstItem.Text
            End If
        Next
        If cllUser.Count = 0 And (bytType = 0 Or bytType = 2) Then
            MsgBox "δѡ��ָ����û���,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        lng������ = Val(lblInfor.Tag)
        If RegRestore(bytType, lng������, cllUser) = False Then
            Exit Function
        End If
        MsgBox "�ָ��ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
        DataBaseToRestoreReg = True
        Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SaveUPData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------
    '����:�ϴ����ݱ���
    '����:
    '����:�ɹ�����true,���򷵻�false
    '---------------------------------------------------------------------------------------------------------------------
    SaveUPData = False
   '�ж��������ֵ�Ƿ���ȷ
    Dim bytType As Byte
    Dim strSQL As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    For i = 1 To txtEdit.UBound
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If i <= 2 Then
            If strTemp = "" Then
                MsgBox txtEdit(i).Tag & "��������!", vbInformation + vbDefaultButton1, gstrSysName
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        If zlCommFun.ActualLen(strTemp) > txtEdit(i).MaxLength Then
            MsgBox txtEdit(i).Tag & "�������,����С�ڵ���" & txtEdit(i).MaxLength & "������" & Int(txtEdit(i).MaxLength / 2) & "������!", vbInformation + vbDefaultButton1, gstrSysName
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, strTemp, "'") > 0 Then
            MsgBox txtEdit(i).Tag & "�����˷Ƿ��ַ���'��!", vbInformation + vbDefaultButton1, gstrSysName
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
    Next
    bytType = Val(Split(cbo��ʽ.Text, "-")(0))
    If bytType = 0 Or bytType = 2 Then
        '��Ҫָ������û�
        If cbo�û�.Text = "" Or cbo�û�.ListIndex < 0 Then
            MsgBox "δָ���ϴ��û���������ָ��!", vbInformation + vbDefaultButton1, gstrSysName
            If cbo�û�.Enabled Then cbo�û�.SetFocus
            Exit Function
        End If
        If InStr(1, mstrPrvs, "�����ϴ�") = 0 Then
            If cbo�û�.Text <> UCase(gstrDbUser) Then
                MsgBox "��û��Ȩ���ϴ�" & cbo�û�.Text & "�û���" & vbCrLf & "���˽�в�����������ָ��!", vbInformation + vbDefaultButton1, gstrSysName
                If cbo�û�.Enabled Then cbo�û�.SetFocus
                Exit Function
            End If
        End If
    End If
    strSQL = "Select ������ ,�û��� from zlClientScheme where ������=[1]"
    Set rsTemp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, Val(txtEdit(1).Text))
    If Not rsTemp.EOF Then
                
        '�ж��û����Ƿ���ͬ
        If zlCommFun.Nvl(rsTemp!�û���) <> UCase(gstrDbUser) And InStr(1, mstrPrvs, "�����ϴ�") = 0 Then
            '�ò������Ѿ����������ã�ϵͳ��Ĭ��һ�¸��º�!"
            MsgBox "�ò������Ѿ����������ã�ϵͳ��Ĭ��һ���º�!", vbInformation + vbDefaultButton1, gstrSysName
            txtEdit(1).Text = Max������
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        Else
            If MsgBox("�÷������Ѿ�����,�Ƿ񸲸Ǹ÷���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtEdit(1).Text = Max������
                If txtEdit(1).Enabled Then txtEdit(1).SetFocus
                Exit Function
            End If
        End If
    End If
    
    Dim cllData As New Collection
    If ExportParasToCollection(bytType, UCase(cbo�û�.Text), cllData) = False Then Exit Function
    If cllData Is Nothing Then
        MsgBox "�����������Ĳ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    If cllData.Count = 0 Then
        MsgBox "�����������Ĳ�����Ϣ,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo ErrHand:
    Call SaveClientParaToDataBase(cllData)
    MsgBox "�ϴ��ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
    SaveUPData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function
Private Sub SaveClientParaToDataBase(ByVal cllData As Collection)
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ͻ��˲������浽���ݿ���
    '����:cllData-������
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    Dim lng������ As Long
    Dim str������ As String
    Dim str����˵�� As String
    Dim strKey As String
    Dim strTemp As String
    Dim strData As String
    lng������ = Val(txtEdit(1).Text)
    str������ = Trim(txtEdit(2).Text)
    str����˵�� = Trim(txtEdit(3).Text)
    Dim rsTemp As New Recordset
    Dim lngĿ¼ As Long, lng���� As Long, lng��ֵ As Long

    strSQL = "select Ŀ¼,����,��ֵ from zlClientparaList where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, strSQL, "ȡ�ֶγ���"
    lngĿ¼ = rsTemp.Fields("Ŀ¼").DefinedSize
    lng���� = rsTemp.Fields("����").DefinedSize
    lng��ֵ = rsTemp.Fields("��ֵ").DefinedSize

    gcnOracle.BeginTrans
    '��ɾ���÷���
    strSQL = "delete zlClientScheme where ������=" & lng������
    gcnOracle.Execute strSQL
    
    strSQL = " insert into zlClientScheme(������,��������, ��������,����վ,�û���) values(" & lng������ & ","
    strSQL = strSQL & "'" & str������ & "',"
    strSQL = strSQL & "'" & str����˵�� & "',"
    strSQL = strSQL & "'" & AnalyseComputer & "',"
    strSQL = strSQL & "'" & UCase(gstrDbUser) & "')"
    gcnOracle.Execute strSQL
    For i = 1 To cllData.Count
        '���ϴ�����
        strSQL = "insert into zlClientparaList(������,���,���,Ŀ¼,����,��ֵ,������Դ,����˵��) values ("
        strSQL = strSQL & "" & lng������ & ","
        strSQL = strSQL & "" & i & ","
        strKey = cllData(i)(0)
        '    If InStr(1, strSection, "˽��ģ��") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "S0" & i
        '    ElseIf InStr(1, strSection, "˽��ȫ��") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "S1" & i
        '    ElseIf InStr(1, strSection, "����ģ��") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "G0" & i
        '    ElseIf InStr(1, strSection, "����ȫ��") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "G1" & i
        '    End If
        If strKey = "˽��ģ��" Then
            strSQL = strSQL & "'˽��ģ��',"
            strTemp = Replace(cllData(i)(1), "˽��ģ��\" & UCase(Trim(cbo�û�.Text)) & "\", "")
            strTemp = Replace(strTemp, "˽��ģ��\" & UCase(Trim(cbo�û�.Text)), "")
        ElseIf strKey = "˽��ȫ��" Then
            strSQL = strSQL & "'˽��ȫ��',"
            strTemp = Replace(cllData(i)(1), "˽��ȫ��\" & UCase(Trim(cbo�û�.Text)) & "\", "")
            strTemp = Replace(strTemp, "˽��ȫ��\" & UCase(Trim(cbo�û�.Text)), "")
        ElseIf strKey = "����ģ��" Then
            strSQL = strSQL & "'����ģ��',"
            strTemp = Replace(cllData(i)(1), "����ģ��\", "")
            strTemp = Replace(strTemp, "����ģ��", "")
        Else
            strSQL = strSQL & "'����ȫ��',"
            strTemp = Replace(cllData(i)(1), "����ȫ��\", "")
            strTemp = Replace(strTemp, "����ȫ��", "")
        End If
        
        strSQL = strSQL & "'" & strTemp & "',"
        strData = Replace(cllData(i)(3), "'", "''")             '��ֵ
        strData = Replace(strData, "\\", "\")
              
        strSQL = strSQL & "'" & cllData(i)(2) & "',"
        strSQL = strSQL & "'" & strData & "',"
        strSQL = strSQL & "0,NULL)"
        If zlCommFun.ActualLen(strTemp) > lngĿ¼ Or zlCommFun.ActualLen(cllData(i)(2)) > lng���� Or zlCommFun.ActualLen(strData) > lng��ֵ Then
            '�������ݿ�Ĵ洢��Χ,�Ͳ�������
            Debug.Print "fds"
        Else
            gcnOracle.Execute strSQL
        End If
    Next
    gcnOracle.CommitTrans
End Sub

Private Sub initData()
    '--------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '--------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset

    
    '--��ʼ���ϴ���ʽ
    Me.cbo��ʽ.Clear
    Me.cbo��ʽ.AddItem "0-���в���(����+˽��)"
    Me.cbo��ʽ.AddItem "1-��������(����ȫ��+����ģ��)"
    Me.cbo��ʽ.AddItem "2-˽�в���(˽��ȫ��+˽��ģ��)"
    Me.cbo��ʽ.ListIndex = 0
    Me.cbo����.Clear
    Me.cbo����.AddItem "0-���в���(����+˽��)"
    Me.cbo����.AddItem "1-��������(����ȫ��+����ģ��)"
    Me.cbo����.AddItem "2-˽�в���(˽��ȫ��+˽��ģ��)"
    Me.cbo����.ListIndex = 0
    
    If mbytParaType = 0 Then
        '��ʼ���ϴ����û���
        strSQL = "Select distinct upper(�û���) �û��� From �ϻ���Ա�� "
        zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
        With rsTemp
            Me.cbo�û�.Clear
            Do While Not .EOF
                Me.cbo�û�.AddItem zlCommFun.Nvl(!�û���)
                If zlCommFun.Nvl(!�û���) = gstrDbUser Then
                    Me.cbo�û�.ListIndex = Me.cbo�û�.NewIndex
                End If
                .MoveNext
            Loop
            If Me.cbo�û�.ListCount = 0 Or Me.cbo�û�.ListIndex < 0 Then
                If gstrDbUser <> "" Then
                    '���뵱ǰ�û�
                    Me.cbo�û�.AddItem gstrDbUser
                    Me.cbo�û�.ListIndex = Me.cbo�û�.NewIndex
                End If
            End If
        End With
        
        '���ҳ���ǰ�û������һ���ƶ��ķ�����
        strSQL = "Select ������,��������, ��������,����վ from zlClientScheme where ������ =(Select max(������) from zlClientScheme where �û��� =[1])"
        Set rsTemp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, gstrDbUser)
        If Not rsTemp.EOF Then
            '��������
            txtEdit(1).Text = zlCommFun.Nvl(rsTemp!������)
            txtEdit(2).Text = zlCommFun.Nvl(rsTemp!��������)
            txtEdit(2).Tag = zlCommFun.Nvl(rsTemp!������)
            txtEdit(3).Text = zlCommFun.Nvl(rsTemp!��������)
        Else
            txtEdit(1).Text = Max������
        End If
        Exit Sub
    End If
    If mbytParaType = 2 Then
        Exit Sub
    End If
    lvw�û�.ListItems.Clear
    Dim strComputerName As String
    
    strComputerName = AnalyseComputer
    strSQL = "Select distinct a.������,a.����վ  ,a.�û���,b.��������,b.��������,b.����վ as �ϴ�վ��,b.�û��� as �ϴ��û���" & _
            " From Zlclientparaset a,zlClientScheme b" & _
            " where a.������=b.������ and (a.����վ=[1] or (a.����վ is null and a.�û��� is not null))"
            
    Set rsTemp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, strComputerName)
    If rsTemp.EOF Then
        MsgBox "��վ��û�����κ���Ҫ�ָ��Ĳ���������" & vbCrLf & "���ܽ��лָ�[�������������еġ��������á�]��", vbInformation + vbDefaultButton1, gstrSysName
        Unload Me
        Exit Sub
    End If
    Dim lst As ListItem
    Dim bln���� As Boolean '���ڹ�������
    bln���� = False
    lblInfor.Caption = "�ָ�������[" & zlCommFun.Nvl(rsTemp!������) & "]" & zlCommFun.Nvl(rsTemp!��������)
    lblInforLst.Caption = "������Ϣ���ϴ�վ��[" & zlCommFun.Nvl(rsTemp!�ϴ�վ��) & "],�ϴ��û�[" & zlCommFun.Nvl(rsTemp!�ϴ��û���) & "]"
    lblInfor.Tag = zlCommFun.Nvl(rsTemp!������)
    Err = 0: On Error Resume Next
    With rsTemp
        Do While Not .EOF
            If Not IsNull(!����վ) And IsNull(!�û���) Then
                '���ڹ��ò���
                bln���� = True
            End If
            If zlCommFun.Nvl(rsTemp!�û���) <> "" Then
                If InStr(1, mstrPrvs, "��������") <> 0 Then
                    Set lst = lvw�û�.ListItems.Add(, "K" & zlCommFun.Nvl(rsTemp!�û���), zlCommFun.Nvl(rsTemp!�û���), "User", "User")
                ElseIf zlCommFun.Nvl(rsTemp!�û���) = UCase(gstrDbUser) Then
                    Set lst = lvw�û�.ListItems.Add(, "K" & zlCommFun.Nvl(rsTemp!�û���), zlCommFun.Nvl(rsTemp!�û���), "User", "User")
                End If
                lst.Checked = True
            End If
            .MoveNext
        Loop
    End With
    If bln���� = False And lvw�û�.ListItems.Count = 0 Then
        MsgBox "��վ�㲻�����κλָ��������������ܽ��лָ���", vbInformation + vbDefaultButton1, gstrSysName
        Unload Me
        Exit Sub
    End If
    Dim i As Long
    For i = 0 To cbo����.ListCount - 1
        Select Case Split(cbo����.List(i), "-")(0)
        Case 0 '����
            If bln���� = False Or lvw�û�.ListItems.Count = 0 Then
                cbo����.RemoveItem i
            End If
        Case 1 '����
            If bln���� = False Then
                cbo����.RemoveItem i
            End If
        Case 2 '˽��
            If lvw�û�.ListItems.Count = 0 Then
                cbo����.RemoveItem i
            End If
        End Select
    Next
    If cbo����.ListCount = 0 Then
        MsgBox "��վ�㲻�����κλָ��������������ܽ��лָ���", vbInformation + vbDefaultButton1, gstrSysName
        Unload Me
        Exit Sub
    End If
    cbo����.ListIndex = 0
End Sub
Private Function Max������() As Long
    '����:��ȡ��������
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    strSQL = "Select nvl(Max(������),0)+1 as ������ from zlClientScheme"
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    If rsTemp.EOF Then
        Max������ = 1
    Else
        Max������ = Val(zlCommFun.Nvl(rsTemp!������))
    End If
End Function

Private Sub cmdSearch_Click()
    Dim strFile As String
    
    Err = 0
    On Error Resume Next
    With Dlg
        .Filter = "ע���ļ�(*.reg)|*.reg"
        .Flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Sub
        strFile = .FileName
    End With
    Err = 0
    txtFile.Text = strFile
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdSel_Click()
        Call SelectScreme("")
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    '��ʼ������
    Call initData
    Call setCtlShowMode
    If mbytParaType = 0 Then
        Me.Caption = "�����ϴ�"
        If Me.txtEdit(1).Enabled And Me.txtEdit(1).Visible Then Me.txtEdit(1).SetFocus
    ElseIf mbytParaType = 1 Then
        Me.Caption = "�����ָ�"
        If Me.cbo����.Enabled And Me.cbo����.Visible Then Me.cbo����.SetFocus
    Else
        Me.Caption = "����ע����Ϣ������ָ�"
        Me.cmdSave.Caption = "�˳�(&X)"
        Me.cmdCancel.Visible = False
        If Me.txtFile.Enabled And Me.txtFile.Visible Then Me.txtFile.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub lvw�û�_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Me.chkAllUser.Value = 2
End Sub

Private Sub lvw�û�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    Call SetSaveCtlEnable
    If Index = 2 Then
        txtEdit(Index).Tag = ""
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
'    If Index <> 1 Then
'        '�����뷨
'        ImeLanguage True
'    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 3 Then
            'ȡ���س���
            KeyCode = 0
        Else
            If Index = 2 And Trim(txtEdit(Index)) <> "" And Trim(txtEdit(2).Tag) = "" Then
                Call SelectScreme(Trim(txtEdit(Index)))
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m����ʽ
    Else
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    End If
    If KeyAscii = vbKeyReturn And Index = 3 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub SetSaveCtlEnable()
    '����:����Save�ؼ���Enable����
    If mbytParaType = 0 Then
        Me.cmdSave.Enabled = Trim(txtEdit(1).Text) <> "" And Trim(txtEdit(2).Text) <> ""
    Else
        Me.cmdSave.Enabled = True
    End If
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub SelectScreme(ByVal strKey As String)
    '����:ѡ�񷽰�
    '����:  strKey -��������
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lstItem As ListItem
    Err = 0: On Error GoTo ErrHand:
    strSQL = "" & _
        "   Select  ������,��������,��������,����վ as �ϴ�վ��,�û��� as �ϴ��û���" & _
        " From zlClientScheme"
    If InStr(1, mstrPrvs, "�����ϴ�") = 0 Then
            strSQL = strSQL & " where �û���='" & UCase(gstrDbUser) & "'"
            If strKey <> "" Then
                strSQL = strSQL & " and ( ������ like '" & strKey & "%' or �������� like '" & strKey & "%' or �û��� like '" & strKey & "%')"
            End If
    Else
        If strKey <> "" Then
            strSQL = strSQL & " where ������ like '" & strKey & "%' or �������� like '" & strKey & "%' or �û��� like '" & strKey & "%'"
        End If
    End If
    strSQL = strSQL & " order by ������"
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "�����ڷ��������ķ���!", vbInformation + vbDefaultButton1, gstrSysName
            If txtEdit(2).Enabled Then txtEdit(2).SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            txtEdit(1).Text = zlCommFun.Nvl(!������)
            txtEdit(2).Text = zlCommFun.Nvl(!��������)
            txtEdit(2).Tag = zlCommFun.Nvl(!������)
            txtEdit(3).Text = zlCommFun.Nvl(!��������)
            If cbo��ʽ.Enabled And fra(0).Visible Then cbo��ʽ.SetFocus
            Exit Sub
        End If
        Me.lvw������.ListItems.Clear
        Do While Not .EOF
            Set lstItem = lvw������.ListItems.Add(, "K" & zlCommFun.Nvl(!������), zlCommFun.Nvl(!������), "Scheame", "Scheame")
            lstItem.SubItems(1) = zlCommFun.Nvl(!��������)
            lstItem.SubItems(2) = zlCommFun.Nvl(!�ϴ��û���)
            lstItem.SubItems(3) = zlCommFun.Nvl(!�ϴ�վ��)
            lstItem.SubItems(4) = zlCommFun.Nvl(!��������)
            If .AbsolutePosition = 1 Then lstItem.Selected = True
            .MoveNext
        Loop
    End With
    With lvw������
        .Top = fra(0).Top + txtEdit(2).Top + txtEdit(2).Height
        .Left = fra(0).Left + txtEdit(2).Left
        .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub lvw������_DblClick()
    Call lvw������_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvw������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvw������.ListItems.Count = 0 Then Exit Sub
    If lvw������.SelectedItem Is Nothing Then Exit Sub
    
    txtEdit(1).Text = lvw������.SelectedItem.Text
    txtEdit(2).Text = lvw������.SelectedItem.SubItems(1)
    txtEdit(2).Tag = lvw������.SelectedItem.Text
    txtEdit(3).Text = lvw������.SelectedItem.SubItems(4)
    
    lvw������.Visible = False
    If cbo��ʽ.Enabled And fra(0).Visible Then cbo��ʽ.SetFocus
End Sub
Private Sub lvw������_LostFocus()
    lvw������.Visible = False
End Sub


