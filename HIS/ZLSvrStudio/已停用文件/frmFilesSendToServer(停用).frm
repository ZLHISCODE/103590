VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilesSendToServer 
   BackColor       =   &H80000005&
   Caption         =   "վ���ļ��ռ�"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmFilesSendToServer.frx":0000
   ScaleHeight     =   6705
   ScaleMode       =   0  'User
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   4080
      TabIndex        =   19
      Top             =   2265
      Width           =   2000
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   6975
      TabIndex        =   17
      Text            =   "21"
      Top             =   1890
      Width           =   420
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H80000005&
      Caption         =   "�ļ�����"
      Height          =   180
      Index           =   0
      Left            =   1245
      TabIndex        =   9
      Top             =   1260
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H80000005&
      Caption         =   "FTP"
      Height          =   180
      Index           =   1
      Left            =   2310
      TabIndex        =   10
      Top             =   1260
      Width           =   810
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "��������(&O)"
      Height          =   300
      Left            =   7500
      TabIndex        =   12
      Top             =   1215
      Width           =   1150
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "��"
      Height          =   290
      Left            =   6795
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1560
      Width           =   285
   End
   Begin VB.Frame fra 
      Height          =   30
      Left            =   -60
      TabIndex        =   8
      Top             =   1140
      Width           =   20000
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   4695
      TabIndex        =   11
      Text            =   "Log;Doc"
      ToolTipText     =   "����ļ�������;�ָ�"
      Top             =   1215
      Width           =   2700
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   16
      Top             =   1890
      Width           =   2000
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   1245
      TabIndex        =   15
      Top             =   1890
      Width           =   2000
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   1245
      TabIndex        =   13
      Top             =   1545
      Width           =   6150
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   300
      Left            =   6120
      TabIndex        =   20
      Top             =   2250
      Width           =   1275
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   3615
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesSendToServer.frx":04F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwClients 
      Height          =   3735
      Left            =   300
      TabIndex        =   2
      Top             =   2580
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilsIcon"
      SmallIcons      =   "ilsIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����վ"
         Object.Tag             =   "����վ"
         Text            =   "����վ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "IP"
         Object.Tag             =   "IP"
         Text            =   "IP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "CPU"
         Object.Tag             =   "CPU"
         Text            =   "CPU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "�ڴ�"
         Object.Tag             =   "�ڴ�"
         Text            =   "�ڴ�"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Ӳ��"
         Object.Tag             =   "Ӳ��"
         Text            =   "Ӳ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "����ϵͳ"
         Object.Tag             =   "����ϵͳ"
         Text            =   "����ϵͳ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "����"
         Object.Tag             =   "����"
         Text            =   "����"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "��;"
         Object.Tag             =   "��;"
         Text            =   "��;"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "˵��"
         Object.Tag             =   "˵��"
         Text            =   "˵��"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CheckBox chkAllSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ȫѡ(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7740
      TabIndex        =   22
      Top             =   2325
      Width           =   1110
   End
   Begin VB.Label lblSource 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ռ���ʽ"
      Height          =   180
      Index           =   5
      Left            =   480
      TabIndex        =   23
      Tag             =   "Ŀ��·��"
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&Z)"
      Height          =   180
      Left            =   3405
      TabIndex        =   18
      Top             =   2295
      Width           =   630
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ʶ˿�"
      Height          =   180
      Index           =   4
      Left            =   6135
      TabIndex        =   21
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label lblSource 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�ռ��ļ�����"
      Height          =   180
      Index           =   3
      Left            =   3585
      TabIndex        =   7
      Tag             =   "Ŀ��·��"
      Top             =   1275
      Width           =   1080
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Index           =   2
      Left            =   3315
      TabIndex        =   6
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label lblSource 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����û���"
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Top             =   1965
      Width           =   900
   End
   Begin VB.Label lblList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ռ��ļ�վ���嵥"
      Height          =   180
      Left            =   315
      TabIndex        =   1
      Top             =   2295
      Width           =   1440
   End
   Begin VB.Label lblSource 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ��·��"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Tag             =   "Ŀ��·��"
      Top             =   1605
      Width           =   720
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "��ָ��վ�������ļ��ռ����ļ�������ָ��Ŀ¼��,��Ŀ¼���ļ���������վ��Ļ�����_վ���������ļ�����"
      Height          =   525
      Left            =   885
      TabIndex        =   4
      Top             =   690
      Width           =   6525
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "վ���ļ��ռ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   3
      Top             =   105
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmFilesSendToServer.frx":0FC3
      Top             =   585
      Width           =   480
   End
End
Attribute VB_Name = "frmFilesSendToServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mintColumn As Integer
Dim mblnChange As Boolean
Dim mintCount As Integer        '��¼��һ�β��ҵ���λ��

Private mintUpType      As Integer  '0 ����ʽ 1 FTP��ʽ'

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

Private Sub chkAllSel_Click()
    Dim itm As ListItem
    If chkAllSel.Tag = "T" Then chkAllSel.Tag = "": Exit Sub
    err = 0
    On Error Resume Next
    Call ExecuteProcedure("Zl_Zlclients_Control(5,Null,Null,Null,Null,Null,Null,Null," & IIf(Me.chkAllSel.value = 1, 1, 0) & ")", Me.Caption)
    For Each itm In Me.lvwClients.ListItems
        itm.Checked = IIf(Me.chkAllSel.value = 1, True, False)
    Next
End Sub
Private Sub cmdRefresh_Click()
    '��ʼ����Ϣ
    Call InitInfor
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Call InitUpType
    '�������ʼ��
    txtFind.Text = "������IP��ַ����վ": txtFind.ForeColor = vbGrayText: mintCount = 0
    '��ʼ����Ϣ
    Call InitInfor
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mintUpType = 0 Then
        txtEdit(4).Visible = False
        lblSource(4).Visible = False
        cmdPath.Caption = "��"
        cmdPath.Width = 285
        cmdPath.Left = 7090
    Else
        txtEdit(4).Visible = True
        txtEdit(4).Text = 21
        lblSource(4).Visible = True
        cmdPath.Caption = "����"
        cmdPath.Width = 615
        cmdPath.Left = 6760
    End If
        
    With lvwClients
        .Width = ScaleWidth - .Left - 50
        .Height = ScaleHeight - 50 - .Top - 50
    End With
    
End Sub
Private Sub SetCtlEnabled()
    Dim blnNoClients As Boolean 'û��վ��
    blnNoClients = Me.lvwClients.ListItems.Count = 0
    chkAllSel.Enabled = Not blnNoClients
End Sub

Private Sub InitInfor()
    '---------------------------------------------------------------------------------------------
    '���ܣ���ʼ����ֵ
    '������
    '���أ�
    '---------------------------------------------------------------------------------------------
    Dim RsFileDirectory As New ADODB.Recordset
    Dim strSQL As String
    Dim bln�ռ�Ŀ¼ As Boolean
    On Error GoTo errHandle
    Set RsFileDirectory = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Reginfo", "")
    With RsFileDirectory
        
        If mintUpType = 0 Then
            Do While Not .EOF
                Select Case IIf(IsNull(!��Ŀ), "", !��Ŀ)
                Case "�ռ�Ŀ¼S"
                    txtEdit(0).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                Case "�����û�S"
                    txtEdit(1).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                Case "��������S"
                    txtEdit(2).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                Case "�ռ�����"
                    txtEdit(3).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                End Select
                .MoveNext
            Loop
        Else
            Do While Not .EOF
                Select Case IIf(IsNull(!��Ŀ), "", !��Ŀ)
                Case "�ռ�Ŀ¼F"
                    txtEdit(0).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                Case "�����û�F"
                    txtEdit(1).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                Case "��������F"
                    txtEdit(2).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                Case "���ʶ˿�F"
                    txtEdit(4).Text = IIf(IsNull(!����), "", !����)
                Case "�ռ�����"
                    txtEdit(3).Text = IIf(IsNull(!����), "", !����)
                    bln�ռ�Ŀ¼ = True
                End Select
                .MoveNext
            Loop
        End If
        
        If bln�ռ�Ŀ¼ = False Then
            MsgBox "ϵͳδ���ڡ��ļ��ռ�Ŀ¼�������ϵͳ����Ա", vbInformation + vbDefaultButton1, gstrSysName
        End If
    End With
    mblnChange = False
    '����վ����Ϣ
    Call LoadClientsInfor
    SetCmd
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub
Private Sub LoadClientsInfor()
    '---------------------------------------------------------------------------------------------
    '���ܣ�����վ����Ϣ
    '������
    '���أ�
    '---------------------------------------------------------------------------------------------
    Dim RsClients As New ADODB.Recordset
    Dim strSQL As String
    Dim itm As ListItem
    
    err = 0
    On Error GoTo errHand:
    Set RsClients = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", "")
    With RsClients

        lvwClients.ListItems.Clear
        lvwClients.Tag = ""
        Do While Not .EOF
            Set itm = lvwClients.ListItems.Add(, "K" & IIf(IsNull(!����վ), "", !����վ), IIf(IsNull(!����վ), "", !����վ), 1, 1)
            itm.SubItems(1) = IIf(IsNull(!IP), "", !IP)
            itm.SubItems(2) = IIf(IsNull(!cpu), "", !cpu)
            itm.SubItems(3) = IIf(IsNull(!�ڴ�), "", !�ڴ�)
            itm.SubItems(4) = IIf(IsNull(!Ӳ��), "", !Ӳ��)
            itm.SubItems(5) = IIf(IsNull(!����ϵͳ), "", !����ϵͳ)
            itm.SubItems(6) = IIf(IsNull(!����), "", !����)
            itm.SubItems(7) = IIf(IsNull(!��;), "", !��;)
            itm.SubItems(8) = IIf(IsNull(!˵��), "", !˵��)
            If !�ռ���־ = 1 Then
                itm.Checked = True
            End If
            .MoveNext
        Loop
    End With

    If Not lvwClients.SelectedItem Is Nothing Then
        lvwClients.SelectedItem.Selected = True
        lvwClients.SelectedItem.EnsureVisible
        lvwClients_ItemClick lvwClients.SelectedItem
    End If
    SetCtlEnabled
    Exit Sub
errHand:
    MsgBox "ϵͳ���ִ���,����Ϊ:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    SetCtlEnabled
    Exit Sub
End Sub

Private Sub lvwClients_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then
        lvwClients.SortOrder = IIf(lvwClients.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwClients.SortKey = mintColumn
        lvwClients.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwClients_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = False Then
        If chkAllSel.value = 1 Then
            chkAllSel.Tag = "T"
            chkAllSel.value = 2
        End If
    End If
    err = 0
    mblnChange = True
    SetCmd
    On Error Resume Next
    Call ExecuteProcedure("Zl_Zlclients_Control(5,'" & UCase(Item.Text) & "',Null,Null,Null,Null,Null,Null," & IIf(Item.Checked, 1, 0) & ")", Me.Caption)
End Sub

Private Sub cmdPath_Click()
    Dim strFolderName As String
    If mintUpType = 0 Then
        strFolderName = OpenFolder(Me, "ѡ���ļ���Ŀ��·��", gstrAPIPath)
        If strFolderName = "" Then Exit Sub
        If Len(strFolderName) = 3 Then
            MsgBox "����ѡ���Ŀ¼(" & strFolderName & ")!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        err = 0
        gstrAPIPath = Trim(strFolderName)
        txtEdit(0).Text = Trim(strFolderName)

        If InStr(1, strFolderName, "\\") <> 0 Then
            Me.txtEdit(0).Text = strFolderName
        Else
            Me.txtEdit(0).Text = "\\" & GetMyCompterName & Mid(strFolderName, 3)
        End If
    Else
        'FTP����
        Call FtpTest
    End If
End Sub
Private Function SaveData() As Boolean
    Dim strSQL As String
    
    SaveData = False
    err = 0
    On Error GoTo errHand:
    gcnOracle.BeginTrans
    

    If mintUpType = 0 Then
        '��ɾ��
        strSQL = "Delete zlregInfo where (��Ŀ = '�ռ�Ŀ¼S' or ��Ŀ = '�����û�S' or ��Ŀ = '��������S' or ��Ŀ = '�ռ�����') "
        gcnOracle.Execute strSQL
        '�ڲ���
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ռ�Ŀ¼S',Null,'" & Trim(Me.txtEdit(0).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�����û�S',Null,'" & Trim(Me.txtEdit(1).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('��������S',Null,'" & Trim(Me.txtEdit(2).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ռ�����',Null,'" & Trim(Me.txtEdit(3).Text) & "')"
        gcnOracle.Execute strSQL
       
    Else
        '��ɾ��
        strSQL = "Delete zlregInfo where (��Ŀ = '�ռ�Ŀ¼F' or ��Ŀ = '�����û�F' or ��Ŀ = '��������F' or ��Ŀ = '���ʶ˿�F' or ��Ŀ = '�ռ�����') "
        gcnOracle.Execute strSQL
        '�ڲ���
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ռ�Ŀ¼F',Null,'" & Trim(Me.txtEdit(0).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�����û�F',Null,'" & Trim(Me.txtEdit(1).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('��������F',Null,'" & Trim(Me.txtEdit(2).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('���ʶ˿�F',Null,'" & Trim(Me.txtEdit(4).Text) & "')"
        gcnOracle.Execute strSQL
        strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ռ�����',Null,'" & Trim(Me.txtEdit(3).Text) & "')"
        gcnOracle.Execute strSQL
    End If
    
    gcnOracle.CommitTrans
    
    SaveData = True
    Exit Function
errHand:
    gcnOracle.RollbackTrans
    MsgBox err.Description
End Function
Private Sub cmdSave_Click()
    If IsValid = False Then Exit Sub
    If Not SaveData Then Exit Sub
    Call SaveUpType
    mblnChange = False
    SetCmd
End Sub
Private Sub SetCmd()
    cmdSave.Enabled = mblnChange
End Sub
Private Function IsValid() As Boolean
    '--------------------------------------------------------------------
    '����:��֤���ݵĺϷ���
    '--------------------------------------------------------------------
    IsValid = False
    
     
    If InStr(1, txtEdit(0).Text, "'") <> 0 Then
        MsgBox "ָ��Ŀ¼�в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "�����û�δ����,�����ÿͻ��˵ķ����û���!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    If InStr(1, txtEdit(1).Text, "'") <> 0 Then
        MsgBox "�����û��в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    If InStr(1, txtEdit(2).Text, "'") <> 0 Then
        MsgBox "���������в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Exit Function
    End If
    IsValid = True
End Function


Private Sub lvwClients_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwClients.Tag <> "" Then
        Call SetSelItemBold(lvwClients.ListItems(lvwClients.Tag), False)
    End If
    Call SetSelItemBold(Item, True)
    lvwClients.Tag = Item.Key
End Sub

Private Sub SetSelItemBold(ByVal itm As ListItem, ByVal blnBold As Boolean)
    Dim i As Integer
        
    '���ñ�ѡ�����ɫ
    itm.Bold = blnBold
    For i = 1 To itm.ListSubItems.Count
        itm.ListSubItems(i).Bold = blnBold
    Next
End Sub

Private Sub OptType_Click(Index As Integer)
    If OptType(0).value = True Then
        mintUpType = 0
    Else
        mintUpType = 1
    End If
    Call ClearTxt
    InitInfor
    Call Form_Resize
    mblnChange = True
    SetCmd
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    SetCmd
End Sub

Private Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--����:����ָ�����ļ����ļ��Ƿ����
    '--����: ������ڴ��ļ�ΪTrue,����ΪFlase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Private Sub InitUpType()
'----------------------------------------------------------------------------------------
'����:��ʼ������ʽ��Ϣ
'----------------------------------------------------------------------------------------
    On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    strSQL = " Select ��Ŀ,���� From zlregInfo where ��Ŀ= '�ռ���ʽ'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)

    If rsTemp.EOF = False Then
        strTemp = Nvl(rsTemp!����, "0")
        If strTemp = "1" Then
             OptType(1).value = True
             mintUpType = 1
        Else
             OptType(0).value = True
             mintUpType = 0
        End If
    Else
        OptType(0).value = True
        mintUpType = 0
    End If
    Exit Sub
errH:
    If err Then
        MsgBox "��ʼ��������ʽ����,������Ϣ����:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub FtpTest()
        '����:���Է������Ƿ��ܹ�����
    On Error GoTo errH
    
    If CheckFileServer = False Then Exit Sub
    
    txtEdit(0).Enabled = False
    txtEdit(4).Enabled = False
    txtEdit(1).Enabled = False
    txtEdit(2).Enabled = False
    cmdSave.Enabled = False
    cmdRefresh.Enabled = False
    cmdPath.Enabled = False
    OptType(0).Enabled = False
    OptType(1).Enabled = False
    
    If IsFtpServer(Trim(txtEdit(0).Text), Trim(txtEdit(1)), Trim(txtEdit(2)), Trim(txtEdit(4))) Then
        MsgBox "�ɹ����ӵ�: " & txtEdit(0).Text, vbOKOnly, gstrSysName
        CancelFtpServer
    Else
        MsgBox "����ʧ�ܣ�����FTP������������!", vbInformation, gstrSysName
    End If
    
    txtEdit(0).Enabled = True
    txtEdit(4).Enabled = True
    txtEdit(1).Enabled = True
    txtEdit(2).Enabled = True
    cmdSave.Enabled = True
    cmdRefresh.Enabled = True
    cmdPath.Enabled = True
    OptType(0).Enabled = True
    OptType(1).Enabled = True
    
    Exit Sub
errH:
    If err Then
        lblSource(5).Caption = ""
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Function CheckFileServer() As Boolean
    '-----------------------------------------------------------------------------
    '����:��鵱ǰ��FTP�������Ƿ���ȷ
    '����:��ǰ���ļ��������ĸ�����ȷ,����true,���򷵻�False
    '����:ף��
    '����:2010/12/09
    '-----------------------------------------------------------------------------
    On Error Resume Next
    CheckFileServer = False
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "δ����FTP������,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Function
    End If
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "�����û�δ����,�����÷������û���!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    If InStr(1, txtEdit(1).Text, "'") <> 0 Then
        MsgBox "�����û��в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    If InStr(1, txtEdit(2).Text, "'") <> 0 Then
        MsgBox "���������в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Exit Function
    End If
    If Trim(txtEdit(4).Text) = "" Then
        MsgBox "FTP���ʶ˿�δ����,�����ö˿�!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(4).Enabled Then txtEdit(4).SetFocus
        Exit Function
    End If
    CheckFileServer = True
    Exit Function
End Function

Private Sub SaveUpType()
'----------------------------------------------------------------------------------------
'����:�޸��ռ����ͷ�ʽ��Ϣ
'----------------------------------------------------------------------------------------
    On Error GoTo errH
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim str��Ŀ As String '��Ŀ
    Dim str���� As String '����
    Dim strSQLTemp As String
    str��Ŀ = "�ռ���ʽ"
    If OptType(0).value Then
        str���� = "0"
    Else
        str���� = "1"
    End If
    strSQL = " Select ��Ŀ,���� From zlregInfo where ��Ŀ= '�ռ���ʽ'"
    
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp.EOF = True Then
        strSQLTemp = "insert into zlregInfo(��Ŀ,����) values ('" & str��Ŀ & "','" & str���� & "')"
        gcnOracle.Execute strSQLTemp

    Else
        strSQLTemp = "delete zlRegInfo where ��Ŀ='" & str��Ŀ & "'"
        gcnOracle.Execute strSQLTemp
        strSQLTemp = "insert into zlregInfo(��Ŀ,����) values ('" & str��Ŀ & "','" & str���� & "')"
        gcnOracle.Execute strSQLTemp
    End If
    
    Exit Sub
errH:
    If err Then
        MsgBox "��������������Ϣʱ����,������Ϣ����:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub ClearTxt()
    txtEdit(0).Text = ""
    txtEdit(1).Text = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(4).Text = ""
End Sub

Private Sub txtFind_Change()
    mintCount = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.ForeColor = vbGrayText Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer

    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        txtFind.Text = Replace(txtFind.Text, " ", "")
        With lvwClients
            For intRow = mintCount + 1 To .ListItems.Count
                If InStr(UCase(.ListItems(intRow).Text), UCase(txtFind.Text)) > 0 Or InStr(.ListItems(intRow).SubItems(1), txtFind.Text) > 0 Then
                    mintCount = intRow
                    .ListItems(intRow).Selected = True
                    .ListItems(intRow).EnsureVisible
                    If lvwClients.Tag <> "" Then
                        Call SetSelItemBold(lvwClients.ListItems(lvwClients.Tag), False)
                    End If
                    Call SetSelItemBold(.ListItems(intRow), True)
                    lvwClients.Tag = .ListItems(intRow).Key
                    Exit For
                End If
            Next

            If intRow = (.ListItems.Count + 1) Then
                If mintCount = 0 Then
                    Call MsgBox("δ�ҵ��롰" & txtFind.Text & "��ƥ�����Ŀ������������IP��ַ����վ��", vbInformation, gstrSysName)
                    txtFind.Text = "": txtFind.SetFocus
                Else
                    mintCount = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "������IP��ַ����վ"
        txtFind.ForeColor = vbGrayText
    End If
End Sub
