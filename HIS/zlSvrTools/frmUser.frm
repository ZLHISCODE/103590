VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H80000005&
   Caption         =   "�û���Ȩ����"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUser.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   8550
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   945
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5835
      TabIndex        =   18
      Top             =   3420
      Width           =   5835
   End
   Begin VB.CommandButton cmdWhole 
      Caption         =   "���������û�"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   6015
      TabIndex        =   17
      Top             =   1350
      Width           =   1440
   End
   Begin VB.CommandButton cmdUpdatePWD 
      Caption         =   "�޸�����(&P)"
      Height          =   350
      Left            =   6015
      TabIndex        =   16
      Top             =   3030
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5625
      ScaleHeight     =   285
      ScaleWidth      =   1065
      TabIndex        =   15
      Top             =   968
      Width           =   1065
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6870
      TabIndex        =   14
      Top             =   960
      Width           =   2610
   End
   Begin VB.CommandButton cmdUnDoLock 
      Caption         =   "�û�����(&J)"
      Height          =   350
      Left            =   6015
      TabIndex        =   13
      Top             =   2700
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ���û�(&D)"
      Height          =   350
      Left            =   6015
      TabIndex        =   11
      Top             =   2370
      Width           =   1440
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸��û�(&M)"
      Height          =   350
      Left            =   6015
      TabIndex        =   10
      Top             =   2040
      Width           =   1440
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "�����û�(&A)"
      Height          =   350
      Left            =   6015
      TabIndex        =   12
      Top             =   1695
      Width           =   1440
   End
   Begin VB.Frame fraFuncs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   915
      TabIndex        =   6
      Top             =   5055
      Width           =   6990
      Begin VB.CommandButton cmdWhole 
         Caption         =   "�ָ������û���ɫ"
         Height          =   350
         Index           =   3
         Left            =   2160
         TabIndex        =   9
         Top             =   0
         Width           =   2160
      End
      Begin VB.CommandButton cmdWhole 
         Caption         =   "���������û���ɫ"
         Height          =   350
         Index           =   2
         Left            =   4320
         TabIndex        =   8
         Top             =   0
         Width           =   2160
      End
      Begin VB.CommandButton cmdWhole 
         Caption         =   "�����ϻ���Ա�ָ��û�"
         Enabled         =   0   'False
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2160
      End
   End
   Begin MSComctlLib.ImageList Img��ͼ�� 
      Left            =   165
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":04F9
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":228B
            Key             =   "Role_Dba"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":480D
            Key             =   "Role_User"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":6D8F
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":7A69
            Key             =   "User1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":E2CB
            Key             =   "UserInfor"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":14B2D
            Key             =   "UserLock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgСͼ�� 
      Left            =   135
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1B38F
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1D121
            Key             =   "Role_Dba"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1F6A3
            Key             =   "Role_User"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":21C25
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":228FF
            Key             =   "User1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":29161
            Key             =   "UserInfor"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":2F9C3
            Key             =   "UserLock"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   1185
      Left            =   945
      TabIndex        =   2
      Top             =   3750
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   2090
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "Img��ͼ��"
      SmallIcons      =   "ImgСͼ��"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   2070
      Left            =   945
      TabIndex        =   3
      Top             =   1320
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   3651
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img��ͼ��"
      SmallIcons      =   "ImgСͼ��"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Settlement"
         Text            =   "�û���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��Ա���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��Ա����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��������"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�û�״̬"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmUser.frx":36225
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ��ɫ"
      Height          =   180
      Left            =   945
      TabIndex        =   5
      Top             =   3525
      Width           =   720
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   945
      TabIndex        =   4
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���Ȩ����"
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
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Menu mnuPopuMenu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuMenuSearch 
         Caption         =   "�����Ź���(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuPopuMenuSearch 
         Caption         =   "���û�����(&U)"
         Index           =   1
      End
      Begin VB.Menu mnuPopuMenuSearch 
         Caption         =   "����Ա����(&P)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==ģ�����
'==============================================================
'���˲˵�ö��
Private Enum menuEnum
    ME_���� = 0
    ME_�û� = 1
    ME_��Ա = 2
End Enum
'�û��б��У���1��ʼ����0��û��
Private Enum UserCol
    Col_��Ա��� = 1
    Col_��Ա���� = 2
    Col_�������� = 3
    Col_�û�״̬ = 4
End Enum

Private Enum WholeEnum
    WE_CreateAllUser = 0 '������Ա��Ϊ�û�
    WE_RestoreAllUser = 1 '�ָ������ϻ���Ա
    WE_RecUserRoles = 2 '��¼�����û���ɫ
    WE_RestoreUserRoles = 3 '�ָ������û���ɫ
End Enum
Private mrsSystem As New ADODB.Recordset
Private mstrBakOwner As String '����ϵͳ��ʷ���������ַ���
Private mstrAllSysOwner As String '����ϵͳ������
Private mstr������ As String '���浱ǰϵͳ����������
Private mintColumn As Integer '

Private mbytSearch As Byte      '0-������������,1-���û�����,2-����Ա����
Private mrsUsers As ADODB.Recordset
Private mLastIndex As Long '�ϴ�ѡ�е��û�

Private mobjTip  As clsTipSwap           '������ʾ�����

'==============================================================
'==�����ӿ�
'==============================================================
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As zlPrintLvw
    Dim rsTmp As ADODB.Recordset
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "ϵͳ�û�"
    Set objPrint.Body.objData = lvwUser
    objPrint.UnderAppItems.Add "Ӧ��ϵͳ��" & cboSystem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
'==============================================================
'==�ؼ��¼�
'==============================================================
Private Sub cboSystem_Click()
    Call FillUser
End Sub

Private Sub cmdAdd_Click()
'�����û�
    Dim blnSucced As Boolean
    If frmUserEdit.UserEdit(mstr������) Then
        Set mrsUsers = Nothing
        Call cboSystem_Click
    End If
End Sub

Private Sub CmdDelete_Click()
'ɾ����Ӧ�û�
    Dim strUser As String, intIndex As Integer
    Dim strRemarks As String
        
    If gblnMustRIS And Not gblnRIS And UCase(gstrSTOwner) = UCase(mstr������) Then
        MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    strUser = lvwUser.SelectedItem.Text
    intIndex = lvwUser.SelectedItem.Index
    If UCase(strUser) = "ZLYB" Then
        MsgBox "����һЩ�����û�������ʹ�ñ�����ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    If UCase(strUser) = "ZLDOC" Then
        MsgBox "���������ĵ�������û�������ʹ�ñ�����ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mrsSystem.Filter = "������='" & strUser & "'"
    If mrsSystem.RecordCount > 0 Then
        MsgBox "�û�" & strUser & "��ϵͳ��" & mrsSystem("����") & "���������ߣ�����ɾ����" & _
            vbCrLf & "�����ȷʵҪɾ�����û�����ʹ��װж�������", vbExclamation, gstrSysName
        Exit Sub
    End If
    If MsgBox("��ȷʵҪɾ���û�" & strUser & "��" & vbCrLf & _
        "���Ѹ��û��µ��������ݿ����һ��ɾ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If Mid(lvwUser.SelectedItem.Tag, 3) = "" Then
        If MsgBox("���û����ܲ����㴴����,��ȷʵҪɾ���û�" & strUser & "��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    '��֤��ݲ��������˵��
    strRemarks = "ɾ���û���" & strUser
    If Not CheckAuditStatus("0402", "ɾ���û�", strRemarks) Then Exit Sub
    On Error Resume Next
    MousePointer = 11
    DoEvents
    gcnOracle.Execute "drop user " & strUser & " cascade"
    If err.Number <> 0 Then
        MsgBox "���û����ܲ������㴴���ģ�ɾ��ʧ�ܡ�", vbExclamation, gstrSysName
        err.Clear: MousePointer = 0
        Exit Sub
    End If
    gcnOracle.Execute "delete from " & mstr������ & ".�ϻ���Ա�� where �û���='" & strUser & "'"
    Call ExecuteProcedure("Zl_Zluserroles_Del('" & strUser & "')", Me.Caption)
    
    '������Ҫ������־
    Call SaveAuditLog(3, "ɾ���û�", strUser, strRemarks)
    
    If UCase(gstrSTOwner) = UCase(mstr������) And gblnRIS And gblnMustRIS Then  '�Ǳ�׼���������
        '֪ͨ�������û��Ѿ���ɾ��
        If Not gobjRIS.UserEdit(3, strUser) Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(UserEdit)δ���óɹ�������ϵ����Ա��", vbInformation, gstrSysName
        End If
    End If
    
    lvwUser.ListItems.Remove intIndex
    If lvwUser.ListItems.Count > 0 Then
        If intIndex > lvwUser.ListItems.Count Then intIndex = lvwUser.ListItems.Count
        lvwUser.ListItems(intIndex).Selected = True
        Call lvwUser_ItemClick(lvwUser.ListItems(intIndex))
    End If
    MousePointer = 0
    Call SetEnable
End Sub

Private Sub cmdModify_Click()
    '�޸��û�
    Dim strItem As String, arrTmp As Variant
    
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    If gblnMustRIS And Not gblnRIS And UCase(gstrSTOwner) = UCase(mstr������) Then
        MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If frmUserEdit.UserEdit(mstr������, lvwUser.SelectedItem.Text, strItem) Then
        If strItem = "" Then
            Set mrsUsers = Nothing
            Call cboSystem_Click
        ElseIf mLastIndex > 0 And mLastIndex < lvwUser.ListItems.Count Then
            arrTmp = Split(strItem, "|")
            lvwUser.ListItems(mLastIndex).SubItems(Col_��Ա���) = arrTmp(0)
            lvwUser.ListItems(mLastIndex).SubItems(Col_��Ա����) = arrTmp(1)
            lvwUser.ListItems(mLastIndex).SubItems(Col_��������) = arrTmp(2)
            lvwUser.ListItems(mLastIndex).Selected = True
            Call lvwUser_ItemClick(lvwUser.ListItems(mLastIndex))
        End If
    End If
End Sub

Private Sub cmdUnDoLock_Click()
    '����:���û����н���
    Dim strKey As String, blnLock As Boolean
    
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    blnLock = Mid(lvwUser.SelectedItem.Tag, 1, 1) <> "1"
    strKey = lvwUser.SelectedItem.Key
    If MsgBox("ȷ��Ҫ" & IIf(blnLock, "����", "����") & "�û�����" & Mid(strKey, 2) & "����", vbInformation + vbOKCancel + vbDefaultButton1) = vbCancel Then Exit Sub
    If LockUser(lvwUser.SelectedItem.Text, blnLock) = False Then Exit Sub
    Call FillUser
    err = 0: On Error Resume Next
    lvwUser.ListItems(strKey).Selected = True
    lvwUser.ListItems(strKey).EnsureVisible
    Call lvwUser_ItemClick(lvwUser.ListItems(strKey))
    Call SetEnable
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub cmdUpdatePWD_Click()
    Dim strUserName As String, strPassword As String
    Dim strError As String
    
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
        
    strUserName = lvwUser.SelectedItem.Text
    strPassword = InputBox("�������µ�����", "�޸�" & strUserName & "������", "123")
    
    If strPassword = "" Then Exit Sub
    
    If gobjRegister.UpdateUserPassword(gcnOracle, strUserName, strPassword, True, strError) Then
        MsgBox "�޸�" & strUserName & "������ɹ���", vbInformation + vbOKOnly, "��ʾ"
        '������Ҫ������־
        Call SaveAuditLog(2, "�޸�����", "�޸��û���" & strUserName & "������")
    Else
        MsgBox "�޸�" & strUserName & "������ʧ�ܡ�" & vbCrLf & strError, vbExclamation, "��ʾ"
    End If
    
    If gstrUserName = strUserName Then
        MsgBox "�޸ĵ�ǰ�û�������֮����Ҫ���µ�¼", vbInformation, "��ʾ"
        frmUserLogin.Show 1
        If gcnOracle.State = adStateClosed Then
            End
        End If
    End If
End Sub

Private Sub cmdWhole_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strDept As String, strError As String, strPrompt As String
    Dim strKey As String, strUserName As String
    Dim blnHaveRis As Boolean
    Dim blnMsgRis As Boolean
    Dim i As Long
    
    On Error GoTo errH
    Select Case Index
        Case WE_CreateAllUser  '���������û�(&1)
            If UCase(gstrSTOwner) = UCase(mstr������) Then   '�Ǳ�׼���������
                blnHaveRis = gblnRIS
                If gblnMustRIS And Not gblnRIS Then
                    MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            strDept = frmUserBatCreate.ShowMe(mstr������)
            '1�������Ա�����λΪӢ����ĸ��������Ա�����Ϊ�û���
            '2�������Ա�����λΪ���֣����ԡ�U+��Ա��š���Ϊ�û���
            '3���û�������û���һ�¡�
            If strDept = "" Then Exit Sub
            strSQL = "Select /*+Rule */" & vbNewLine & _
                        " a.Id, a.���, a.����, a.����" & vbNewLine & _
                        "From " & mstr������ & ".��Ա�� a," & mstr������ & ".������Ա b, Table(Cast(f_Num2list('" & strDept & "') As Zltools.t_Numlist)) c" & vbNewLine & _
                        "Where a.Id = b.��Աid And b.����id = c.Column_Value And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And" & vbNewLine & _
                        "      Id Not In (Select ��Աid From " & mstr������ & ".�ϻ���Ա��)" & vbNewLine & _
                        "Order By a.���"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            If rsTmp Is Nothing Then Exit Sub
            On Error Resume Next
            With rsTmp
                Call ShowFlash("���ڴ����û������Ժ�", 0)
                For i = 1 To .RecordCount
                    If UCase(Left(!���, 1)) >= "A" And UCase(Left(!���, 1)) <= "Z" Then
                        strUserName = !���
                    Else
                        strUserName = "U" & !���
                    End If
                    frmMDIMain.stbThis.Panels(2).Text = "���ڴ����û�:" & strUserName
                    Call ShowFlash("���ڴ����û�:��" & strUserName & "��", i / .RecordCount)
                    strError = ""
                    Call gobjRegister.CreateUser(gcnOracle, strUserName, strUserName, strError)
                    If strError = "" Then
                        gcnOracle.Execute "Grant Connect,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to " & strUserName
                        gcnOracle.Execute "insert into " & mstr������ & ".�ϻ���Ա��(�û���,��Աid) values ('" & strUserName & "'," & !id & ")"
                        Call AlterUserTableSpaces(gcnOracle, strUserName)
                        '֪ͨ�������û�������
                        If blnHaveRis Then
                            If Not gobjRIS.UserEdit(1, strUserName) Then
                                blnMsgRis = True
                            End If
                        End If
                    Else
                        strPrompt = strPrompt & vbCrLf & "[" & !strUserName & "]" & !���� & ":" & strError
                    End If
                    .MoveNext
                Next
                Call ShowFlash("")
                If strPrompt = "" Then
                    strPrompt = "ȫ����Ա��ȷ����Ϊ�ϻ��û���"
                Else
                    strPrompt = "������Աδ��������Ϊ�ϻ��û���" & strPrompt
                End If
                '������Ҫ������־
                Call SaveAuditLog(2, "���������û�", strPrompt)
                If blnMsgRis Then
                    strPrompt = strPrompt & vbNewLine & "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(UserEdit)δ���óɹ�������ϵ����Ա��"
                End If
                MsgBox strPrompt, vbInformation, gstrSysName
            End With
            '�����û���ɫδ������������������û��������û���δ��Ӧ��ɫ����˲������û���ɫ��¼
        Case WE_RestoreAllUser '�����ϻ���Ա�ָ��û�(&2)
            If MsgBox("�������û��ڡ����û��ָ����ݡ�ģʽ�£��ָ�����֮�󴴽���ǰ���û�����Ȩ�������û�����Ϊ��ʼ���롣" & vbCrLf _
                    & "��ȷ��Ҫ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            strSQL = "Select �û��� From " & mstr������ & ".�ϻ���Ա�� Where �û��� Not In (Select Username From All_Users)"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            With rsTmp
                On Error Resume Next
                Call ShowFlash("���ڴ����û������Ժ�", 0)
                For i = 1 To .RecordCount
                    frmMDIMain.stbThis.Panels(2).Text = "���ڴ����û�:" & !�û���
                    Call ShowFlash("���ڴ����û���" & !�û��� & "��", i / .RecordCount)
                    Call gobjRegister.CreateUser(gcnOracle, !�û���, !�û���, strError)
                    If strError = "" Then
                        gcnOracle.Execute "Grant Connect,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to " & !�û���
                        Call AlterUserTableSpaces(gcnOracle, Nvl(!�û���))
                    Else
                        strPrompt = strPrompt & vbCrLf & !�û��� & ":" & strError
                    End If
                    .MoveNext
                Next
                Call ShowFlash("")
                If strPrompt = "" Then
                    strPrompt = "�ϻ��û��ָ���ϣ�"
                Else
                    strPrompt = "�����ϻ��û�û�лָ���" & strPrompt
                End If
                MsgBox strPrompt, vbExclamation, gstrSysName
            End With
            '�����û���ɫδ������������������û��������û���δ��Ӧ��ɫ����˲������û���ɫ��¼
            '������Ҫ������־
            Call SaveAuditLog(2, "�����ϻ���Ա�ָ��û�", strPrompt)
        Case WE_RecUserRoles '���������û���ɫ(&3)
            If MsgBox("�����ܽ������ϵͳ����������û��Ľ�ɫ�������û������ݿ���ʵ��ӵ�еĽ�ɫ���²�����ϵͳ�������û��Ľ�ɫ���ݡ�" & vbCrLf & _
                        "���û���Ӧ��ϵͳ�еĽ�ɫ�����ݿ���ʵ�ʵĽ�ɫ��һ��ʱ��ִ�д˲�����������һ�µ����ݡ�" & vbCrLf & _
                        "��ȷ��Ҫ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            If Not CheckRushHours("0402", "���������û���ɫ") Then
                Exit Sub
            End If
            Call ExecuteProcedure("Zl_Zluserroles_Add()", Me.Caption)
            Call FillRole
            MsgBox "���������û���ɫ��������ɡ�", vbInformation, gstrSysName
            '������Ҫ������־
            Call SaveAuditLog(2, "���������û���ɫ", "���������û���ɫ���������")
        Case WE_RestoreUserRoles '�ָ������û���ɫ(&4)
            If MsgBox("�����ܽ����û���Ӧ��ϵͳ�еļ�¼��ɫ���½��н�ɫ��Ȩ��һ�����ڡ����û��ָ����ݡ�ģʽ�£��ָ���ɫ���û�֮���ؽ��û��Ľ�ɫ��" & vbCrLf & _
                        "��ȷ��Ҫ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            If Not CheckRushHours("0402", "�ָ������û���ɫ") Then
                Exit Sub
            End If
            strSQL = "Select �û�, ��ɫ, ���� From Zltools.Zluserroles Where �û� In (Select Username From All_Users)"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            With rsTmp
                On Error Resume Next
                Call ShowFlash("���������û���ɫ�����Ժ�", 0)
                For i = 1 To .RecordCount
                    frmMDIMain.stbThis.Panels(2).Text = "���������û� " & !�û� & " " & !��ɫ
                    Call ShowFlash("���������û���" & !�û� & "�� " & !��ɫ, i / .RecordCount)
                    gcnOracle.Execute "Grant " & !��ɫ & " to " & !�û� & IIf(!���� = 1, " With Admin Option", "")
                    If err.Number <> 0 Then strPrompt = strPrompt & vbCrLf & !��ɫ & "����" & !�û� & "ʧ��": err.Clear
                    .MoveNext
                Next
                Call ShowFlash("")
                If strPrompt = "" Then
                    strPrompt = "�û���ɫ�ָ����"
                Else
                    strPrompt = "�����û���ɫû�лָ���" & strPrompt
                End If
                MsgBox strPrompt, vbExclamation, gstrSysName
                Call FillRole
                frmMDIMain.stbThis.Panels(2).Text = "���������û� " & !�û� & " " & !��ɫ
            End With
            '������Ҫ������־
            Call SaveAuditLog(2, "�ָ������û���ɫ", strPrompt)
    End Select
    
    frmMDIMain.stbThis.Panels(2).Text = ""
    '���¼����û������ָ�ԭʼѡ��
    If Index = WE_CreateAllUser Or Index = WE_RestoreAllUser Then
        If Not lvwUser.SelectedItem Is Nothing Then strKey = lvwUser.SelectedItem.Key
        On Error GoTo errH
        Call FillUser
        err = 0: On Error Resume Next
        lvwUser.ListItems(strKey).Selected = True
        Call lvwUser_ItemClick(lvwUser.ListItems(strKey))
        Call SetEnable
        If err.Number <> 0 Then err.Clear
    End If
    Exit Sub
errH:
    frmMDIMain.stbThis.Panels(2).Text = ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    picHLine.Top = Val(GetSetting("ZLSOFT", "����ģ��\������������\�û�����", "PicHLine_TOP", "6500"))
    gblnMustRIS = Val(gclsBase.GetPara(255, 100, 0, , , "0")) = 1
    If gblnMustRIS Then
        Call CheckAndAdjustMustTable("zlParameters", "����") '�߰汾�������ӵͰ汾���ݿ�����
        gblnRIS = GetRIS
        If gblnRIS Then
            Call gobjRIS.InitConn(gcnOracle)
        End If
    Else
        gblnRIS = False
    End If
    mbytSearch = ME_����: mnuPopuMenuSearch(ME_����).Checked = True: txtSearch.Tag = "�����Ź���"
    Call PrintSearch("�����Ź���", vbBlue, False)
    If gstrSTOwner = "" Then
        gstrSTOwner = GetOwnerName(100, gcnOracle)
    End If
    '�û�״̬��
    lvwUser.ColumnHeaders(Col_�������� + 1).Width = lvwUser.ColumnHeaders(Col_�������� + 1).Width + IIf(gblnDBA, 0, 1000)
    lvwUser.ColumnHeaders(Col_�û�״̬ + 1).Width = IIf(gblnDBA, 1000, 0)
    cmdUnDoLock.Visible = gblnDBA
    cmdUpdatePWD.Visible = gblnDBA
    
    mstrBakOwner = ""
    On Error GoTo errH
    strSQL = "Select Upper(������) ������ From Zlbakspaces Where Db���� Is Null Order by ������"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        If strTmp <> rsTmp!������ Then
            strTmp = rsTmp!������
            mstrBakOwner = mstrBakOwner & ",'" & strTmp & "'"
        End If
        rsTmp.MoveNext
    Loop
    mstrAllSysOwner = ""
    Call FillSystem
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Resize()
    Dim lngTemp As Long
    err = 0: On Error Resume Next
    Me.cmdAdd.Left = Me.ScaleWidth - 200 - Me.cmdAdd.Width
    Me.cmdDelete.Left = Me.cmdAdd.Left
    Me.cmdModify.Left = Me.cmdAdd.Left
    Me.cmdUnDoLock.Left = Me.cmdAdd.Left
    Me.cmdUpdatePWD.Left = Me.cmdAdd.Left
    Me.cmdWhole(WE_CreateAllUser).Left = Me.cmdAdd.Left
    Me.lvwUser.Width = Me.cmdAdd.Left - 90 - Me.lvwUser.Left
    Me.lvwRole.Width = Me.ScaleWidth - Me.lvwRole.Left - 200
    fraFuncs.Top = Me.ScaleHeight - fraFuncs.Height - 100
    picHLine.Width = lvwRole.Width
    lvwUser.Height = picHLine.Top - lvwUser.Top
    lblRole.Top = picHLine.Top + picHLine.Height
    lvwRole.Top = lblRole.Top + lblRole.Height + 50
    lvwRole.Height = fraFuncs.Top - lvwRole.Top - 100
    
    txtSearch.Left = lvwUser.Left + lvwUser.Width - txtSearch.Width
    picSel.Left = txtSearch.Left - picSel.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    SaveSetting "ZLSOFT", "����ģ��\������������\�û�����", "PicHLine_TOP", picHLine.Top
    Set mrsSystem = Nothing
    Set mobjTip = Nothing
    mstr������ = ""
End Sub

Private Sub lvwRole_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    Dim strTip As String, strTitle As String
    
    If mobjTip Is Nothing Then
        Call InitTips
    End If
    Set objItem = lvwRole.HitTest(x, y)
    If Not objItem Is Nothing Then
        If objItem.SubItems(1) = 1 Then
            If glngSysNo <> -1 Then
                strTip = "�ý�ɫ��Ȩ�������⣬�����Զ�ϵͳ��ʽ��¼������" & vbNewLine & "ִ�С����������û���ɫ��������"
            Else
                strTip = "�ý�ɫ��Ȩ�������⣬����ִ�С����������û���ɫ��������"
            End If
            strTitle = objItem.Text
        ElseIf objItem.SubItems(1) = 2 Then
            If glngSysNo <> -1 Then
                strTip = "�ý�ɫ��Ȩ�������⣬�����Զ�ϵͳ��ʽ��¼������" & vbNewLine & "ִ�С��ָ������û���ɫ��������"
            Else
                strTip = "�ý�ɫ��Ȩ�������⣬����ִ�С��ָ������û���ɫ��������"
            End If
            strTitle = objItem.Text
        ElseIf objItem.SubItems(1) = 4 Then
            If glngSysNo <> -1 Then
                strTip = "�ý�ɫ�����ڣ������Զ�ϵͳ��ʽ��¼������" & vbNewLine & "ִ�С����������û���ɫ��������"
            Else
                strTip = "�ý�ɫ�����ڣ�����ִ�С����������û���ɫ��������"
            End If
            strTitle = objItem.Text
        Else
            strTip = ""
            strTitle = ""
        End If
        mobjTip.TipText = strTip
        mobjTip.Title = strTitle
    Else
        mobjTip.TipText = ""
        mobjTip.Title = ""
    End If
End Sub

Private Sub InitTips()
    Set mobjTip = New clsTipSwap
    Set mobjTip.ParentControl = lvwRole
    mobjTip.Icon = TTIconInfo
    mobjTip.Style = TTBalloon
    mobjTip.Create
End Sub

Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwUser.SortOrder = IIf(lvwUser.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwUser.SortKey = mintColumn
        lvwUser.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillRole
    mLastIndex = Item.Index
End Sub

Private Sub mnuPopuMenuSearch_Click(Index As Integer)
    Dim i As Integer
    mbytSearch = Index
    For i = ME_���� To ME_��Ա
        mnuPopuMenuSearch(i).Checked = i = Index
    Next
    txtSearch.Tag = Split(mnuPopuMenuSearch(Index).Caption, "(")(0)
    Call PrintSearch(txtSearch.Tag, vbBlue, False)
    If txtSearch.Enabled Then txtSearch.SetFocus
    If txtSearch.Text <> "" Then
        Call FillUser(True)
    End If
End Sub

Private Sub picHLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHLine.BackColor = &H8000000F: Me.picHLine.Top = Me.picHLine.Top + y
End Sub

Private Sub picHLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picHLine.BackColor = Me.BackColor
    If Me.picHLine.Top < 4000 Then Me.picHLine.Top = 4000
    If Me.picHLine.Top > Me.ScaleHeight - 3000 Then Me.picHLine.Top = Me.ScaleHeight - 3000
    Call Form_Resize
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picSel.Tag = "In" Then
        If x < 0 Or y < 0 Or x > picSel.Width Or y > picSel.Height Then
            ReleaseCapture
            picSel.Tag = ""
            PrintSearch Me.txtSearch.Tag, vbBlue, False
        End If
    Else
        picSel.Tag = "In"
        SetCapture picSel.hwnd
        MousePointer = 99
        PrintSearch Me.txtSearch.Tag, vbRed, True
    End If
End Sub

Private Sub picSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    PopupMenu Me.mnuPopuMenu, vbPopupMenuRightAlign, Me.picSel.Left + 600, Me.picSel.Top + Me.picSel.Height
    Call PrintSearch(Me.txtSearch.Tag, vbBlue, False)
    picSel.Tag = ""
End Sub


Private Sub txtSearch_Change()
    Call FillUser(True)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub


Private Sub PrintSearch(ByVal strTittle As String, ByVal lngColor As Long, ByVal blnBoderStyle As Boolean)
    '----------------------------------------------------------------------------------------
    '����:��ӡָ������������
    '����:strTittle-����
    '     lngColor-��ɫֵ
    '     lngBoderStyl-�Ƿ�ӱ߿���
    '----------------------------------------------------------------------------------------
    '����:��ӡʱ�䷶Χ
    With picSel
        picSel.Width = 980
        .Left = txtSearch.Left - .Width - 50
        .Cls
        '.FontUnderline = blnBoderStyle ' IIf(blnBoderStyle, 1, 0)
        '.ScaleWidth = .TextWidth(strTittle)
        .ForeColor = lngColor
         .FontUnderline = True
        .CurrentX = 10 '(.ScaleWidth - .TextWidth(strTittle))
        .CurrentY = (.ScaleHeight - .TextHeight(strTittle)) / 2
        picSel.Print strTittle
        .ZOrder 1
    End With
End Sub

Private Sub FillSystem()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    
    '�жϸ��û��ܷ񴴽��û�
    On Error GoTo errH
    strSQL = "Select 1" & vbNewLine & _
                    "From User_Sys_Privs" & vbNewLine & _
                    "Where Privilege = 'CREATE USER'" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 1" & vbNewLine & _
                    "From Role_Sys_Privs" & vbNewLine & _
                    "Where Privilege = 'CREATE USER'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cmdAdd.Enabled = rsTmp.RecordCount > 0
    'û��ϵͳʱ���ɾ���޸Ĳ�����
    cmdAdd.Enabled = rsTmp.RecordCount > 0
    cmdDelete.Enabled = cmdAdd.Enabled
    cmdModify.Enabled = cmdAdd.Enabled
    cmdWhole(WE_CreateAllUser).Enabled = cmdAdd.Enabled
    cmdWhole(WE_RestoreAllUser).Enabled = cmdAdd.Enabled
    
    '��ʾϵͳ�����߾��в�����Ա�����ϵͳ
    If glngSysNo <> -1 Then
        cmdWhole(WE_RecUserRoles).Visible = False
        cmdWhole(WE_RestoreUserRoles).Visible = False
        Set mrsSystem = gclsBase.GetMenSystems(True)
        mrsSystem.Filter = "��� = " & glngSysNo
    Else
        Set mrsSystem = gclsBase.GetMenSystems(True, True)
    End If
    If mrsSystem.RecordCount <= 0 Then Exit Sub
    Do While Not mrsSystem.EOF
        If strTmp <> mrsSystem!������ Then
            strTmp = mrsSystem!������
            mstrAllSysOwner = mstrAllSysOwner & "," & strTmp
        End If
        mrsSystem.MoveNext
    Loop
    mstrAllSysOwner = mstrAllSysOwner & ","
    '����ϵͳ����󴥷�ϵͳѡ��
    '��¼�����ˣ���ֵĬ������
    If mrsSystem.RecordCount = 1 Then
        lblSys.Visible = False
        cboSystem.Visible = False
    Else
        mrsSystem.Filter = "��Ա����=1": mrsSystem.Sort = "�����,���"
    End If
    cboSystem.Clear: cboSystem.Tag = ""
    mrsSystem.MoveFirst
    Do While Not mrsSystem.EOF
        cboSystem.addItem mrsSystem!���� & " v" & mrsSystem!�汾�� & "��" & mrsSystem!��� & "��"
        cboSystem.ItemData(cboSystem.NewIndex) = mrsSystem!���
        If mrsSystem!������ & "" = UCase(gstrUserName) And cboSystem.Tag = "" Then
            cboSystem.Tag = cboSystem.NewIndex
        End If
        mrsSystem.MoveNext
    Loop
    cboSystem.ListIndex = Val(cboSystem.Tag) '����Click�¼��������û�
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, Me.Caption
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Sub FillUser(Optional blnFilter As Boolean = False)
'���ܣ�����û�
'blnFilter=�Ƿ����ģʽ
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSearch As String, strIco As String
    Dim lst As ListItem, blnLock As Boolean
    Dim blnOwner As Boolean, bln���� As Boolean
    
    On Error GoTo errH
    '��ʾ���Խ��е�ǰϵͳ���û����Ӧ����Ա
    mrsSystem.Filter = "���=" & cboSystem.ItemData(cboSystem.ListIndex)
    mstr������ = mrsSystem!������
    If blnFilter And Not mrsUsers Is Nothing Then
    Else
        '��ʷ���ݿռ䲻Ӧ�����û�����
        '����ϵͳ�������߲����룬�����������ϵͳ����������Ȩ����Ϊһ�������ߵĶ�����ܺ�����ϵͳ�Ĺ���ͬ��ʳ�ͻ
        If gblnDBA Then
            strSQL = "Select u.Username, ���, ����, ��Ա����, ���ű���, ��������, ���ż���, m.Account_Status" & vbNewLine & _
                            "From All_Users u," & vbNewLine & _
                            "     (Select c.�û���, p.���, p.����, p.���� As ��Ա����, d.���� As ���ű���, d.���� As ��������, d.���� As ���ż���" & vbNewLine & _
                            "       From " & mstr������ & ".��Ա�� p, " & mstr������ & ".���ű� d, " & mstr������ & ".������Ա b, " & mstr������ & ".�ϻ���Ա�� c" & vbNewLine & _
                            "       Where p.Id = c.��Աid And c.��Աid = b.��Աid And d.Id = b.����id And" & vbNewLine & _
                            "             (p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) And b.ȱʡ = 1) r, Dba_Users m" & vbNewLine & _
                            "Where u.Username = r.�û���(+) And u.Username Not In (" & G_STR_USERS & mstrBakOwner & ") And u.User_Id = m.User_Id And" & vbNewLine & _
                            "      Not m.Default_Tablespace In ('SYSTEM', 'DRSYS') And u.Username Not Like 'ZLBAK%' And u.Username Not Like 'ZLHD%'"
        Else
            strSQL = "Select u.Username, ���, ����, ��Ա����, ���ű���, ��������, ���ż���, 'OPEN' As Account_Status" & vbNewLine & _
                            "From All_Users u," & vbNewLine & _
                            "     (Select c.�û���, p.���, p.����, p.���� As ��Ա����, d.���� As ���ű���, d.���� As ��������, d.���� As ���ż���" & vbNewLine & _
                            "       From " & mstr������ & ".��Ա�� p, " & mstr������ & ".���ű� d, " & mstr������ & ".������Ա b, " & mstr������ & ".�ϻ���Ա�� c" & vbNewLine & _
                            "       Where p.Id = c.��Աid And c.��Աid = b.��Աid And d.Id = b.����id And" & vbNewLine & _
                            "             (p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) And b.ȱʡ = 1) r" & vbNewLine & _
                            "Where u.Username = r.�û���(+) And u.Username Not In (" & G_STR_USERS & mstrBakOwner & ") And u.Username Not Like 'ZLBAK%' And u.Username Not Like 'ZLHD%'"
        End If
        Set mrsUsers = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    End If
    
    '���ݹ���
    strSearch = Replace(Trim(UCase(txtSearch.Text)), "'", "")
    lvwUser.ListItems.Clear
    If strSearch = "" Then
        mrsUsers.Filter = 0
    Else
        Select Case mbytSearch
            Case ME_��Ա '����Ա
                mrsUsers.Filter = "��� like '" & strSearch & "%' or ���� like '" & strSearch & "%' or ��Ա���� like '" & strSearch & "%'"
            Case ME_�û� '���û�
                mrsUsers.Filter = "USERNAME like '" & strSearch & "%'"
            Case Else
                '����������
                mrsUsers.Filter = "���ű��� like '" & strSearch & "%' or �������� like '" & strSearch & "%' or ���ż��� like '" & strSearch & "%'"
        End Select
    End If
    '���ݼ���
    With mrsUsers
        Do While Not .EOF
            blnOwner = InStr(mstrAllSysOwner, "," & !USERNAME & ",") > 0
            If Not blnOwner Or gstrUserName = !USERNAME Then
                strIco = "User": blnLock = UCase(!ACCOUNT_STATUS & "") <> "OPEN"
                bln���� = UCase(!ACCOUNT_STATUS & "") = "EXPIRED"
                If blnLock Then
                    strIco = "UserLock"
                ElseIf IsNull(!����) And Not blnOwner Then
                    strIco = "UserInfor"
                End If
                Set lst = lvwUser.ListItems.Add(, "K" & !USERNAME, !USERNAME, strIco, strIco)
                lst.SubItems(Col_��Ա���) = !��� & ""
                lst.SubItems(Col_��Ա����) = !���� & ""
                lst.SubItems(Col_��������) = !�������� & ""
                lst.SubItems(Col_�û�״̬) = IIf(blnLock, IIf(bln����, "�������", "����"), "")
                lst.Tag = IIf(blnLock And Not bln����, "1", "0") & IIf(blnOwner, 1, 0) & !����
            End If
            mrsUsers.MoveNext
        Loop
    End With
    
    If lvwUser.ListItems.Count > 0 Then
        If mLastIndex > 0 And mLastIndex < lvwUser.ListItems.Count Then
            lvwUser.ListItems(mLastIndex).Selected = True
        Else
            lvwUser.ListItems(1).Selected = True
        End If
        Call FillRole
    End If
    Call SetEnable
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillRole()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strUser As String
    Dim lst As ListItem
    
    On Error GoTo errH
    lvwRole.ListItems.Clear
    If lvwUser.SelectedItem Is Nothing Then
        Exit Sub
    Else
        strUser = lvwUser.SelectedItem.Text
    End If
    '��ʾ���û����еĽ�ɫ
    '���Ϊ1��ʾDba_Role_Privs���н�ɫ��Zluserroles��û��
    '���Ϊ2��ʾDba_Role_Privs��û�н�ɫ��Zluserroles���У����ҽ�ɫ�Ǵ��ڵ�
    '���Ϊ3��ʾ�������й���
    '���Ϊ4��ʾDba_Role_Privs��û�н�ɫ��Zluserroles���У����ҽ�ɫ�ǲ����ڵ�
    strSQL = "Select ��ɫ, Sum(���) ���" & vbNewLine & _
            "From (Select Granted_Role ��ɫ, 1 ���" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Grantee = [1] And Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select a.��ɫ, Decode(b.����, Null, 4, 2) ���" & vbNewLine & _
            "       From Zluserroles a, Zlroles b" & vbNewLine & _
            "       Where a.�û� = [1] And a.��ɫ = b.����(+))" & vbNewLine & _
            IIf(glngSysNo = -1, "", " a Where Exists (Select 1 From Zlroles b Where a.��ɫ = b.���� And b.ϵͳ = [2])") & vbNewLine & _
            "Group By ��ɫ"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strUser, cboSystem.ItemData(cboSystem.ListIndex))
    Do While Not rsTmp.EOF
        If rsTmp!��� = 1 Then
            Set lst = lvwRole.ListItems.Add(, , Mid(rsTmp!��ɫ & "", 4), "Role_User")
        ElseIf rsTmp!��� = 2 Or rsTmp!��� = 4 Then
            Set lst = lvwRole.ListItems.Add(, , Mid(rsTmp!��ɫ & "", 4), "Role_Dba")
        Else
            Set lst = lvwRole.ListItems.Add(, , Mid(rsTmp!��ɫ & "", 4), "Role")
        End If
        lst.SubItems(1) = rsTmp!���
        rsTmp.MoveNext
    Loop
    Call SetEnable
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub SetEnable()
'���ø�����ť��Enable����
    Dim blnHave As Boolean
    Dim blnLock As Boolean
    Dim blnOwner As Boolean '������
    
    blnHave = Not lvwUser.SelectedItem Is Nothing
    blnOwner = False
    If blnHave Then
        blnLock = Mid(lvwUser.SelectedItem.Tag, 1, 1) = "1"
        blnOwner = Mid(lvwUser.SelectedItem.Tag, 2, 1) = "1"
    End If
    cmdDelete.Enabled = cmdAdd.Enabled And blnHave And Not blnLock And blnOwner = False
    If cmdDelete.Enabled = True Then
        If lvwUser.SelectedItem.Text = "ZLTOOLS" Then cmdDelete.Enabled = False
    End If
    cmdModify.Enabled = blnHave And Not blnLock
    If blnLock = True Then
        cmdUnDoLock.Caption = "�����û�(&S)"
    Else
        cmdUnDoLock.Caption = "�����û�(&J)"
    End If
End Sub

Private Function LockUser(ByVal strUser As String, Optional ByVal blnLock As Boolean = True) As Boolean
'����:���ָ���û����м��������
'����:strUser-�û���
'     blnLock-true:����;false-����
'�ɹ�:�ӽ����ɹ�,����true,���򷵻�false
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If blnLock Then
        '��Ҫ�ж��Ƿ��������û����������˵�.
        strSQL = "Select Osuser, Machine, Terminal As �ն�, Program From gV$session Where Username = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strUser)
        If Not rsTmp.EOF Then
            If MsgBox("����: " & vbCrLf & "   �û�" & strUser & "�����������ݿ���,���ö��Ѿ���½���û�����Ч,�Ƿ�Ҫ�Ը��û����н���?", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    strSQL = "alter user " & strUser & " ACCOUNT " & IIf(blnLock, "LOCK", "unlock ")
    '�����ͼ���
    err = 0: On Error Resume Next
    gcnOracle.Execute strSQL
    If err.Number <> 0 Then
        MsgBox "����û�[" & strUser & "]��" & IIf(blnLock, "����", "����") & "ʧ��,���Ժ��ټ���!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        err.Clear
        Exit Function
    End If
    '������Ҫ������־
    Call SaveAuditLog(2, "��ͣ�û�", IIf(blnLock, "����", "����") & "�û���" & strUser)
    LockUser = True
End Function

