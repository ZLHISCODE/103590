VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmFilesSet 
   Caption         =   "�ļ���������������"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "frmFilesSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9495
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmd�������� 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8175
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5850
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTP�������� 
      Height          =   300
      Left            =   8175
      TabIndex        =   29
      Top             =   5505
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   119472129
      CurrentDate     =   40908
   End
   Begin VB.CheckBox chk�������� 
      Caption         =   "��������"
      Height          =   240
      Left            =   8175
      TabIndex        =   28
      Top             =   5220
      Width           =   1020
   End
   Begin VB.OptionButton OptType 
      Caption         =   "�ռ�����"
      Height          =   180
      Index           =   1
      Left            =   8310
      TabIndex        =   27
      ToolTipText     =   "�ռ�����ʱ����бȽϲ�����������Ƿ���ͬ,����ͬ�Ľ����ռ�."
      Top             =   3090
      Width           =   1200
   End
   Begin VB.OptionButton OptType 
      Caption         =   "�ռ�����"
      Height          =   180
      Index           =   0
      Left            =   8310
      TabIndex        =   26
      ToolTipText     =   "�ռ�����ʱ�ռ����еĲ���."
      Top             =   2685
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.FileListBox FileList 
      Height          =   630
      Left            =   15
      TabIndex        =   22
      Top             =   5115
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdSaveInfo 
      Caption         =   "��������"
      Height          =   350
      Left            =   8250
      TabIndex        =   16
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "�ռ�(&R)"
      Height          =   350
      Left            =   8250
      TabIndex        =   17
      Top             =   710
      Width           =   1100
   End
   Begin VB.Frame fra�ļ����� 
      Caption         =   "�����ļ��嵥"
      Height          =   3630
      Left            =   75
      TabIndex        =   14
      Top             =   2580
      Width           =   8055
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   3255
         Left            =   195
         TabIndex        =   15
         Top             =   255
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5741
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit mshBillShow 
         Height          =   3255
         Left            =   195
         TabIndex        =   25
         Top             =   225
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5741
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�ϴ�(&O)"
      Height          =   350
      Left            =   8250
      TabIndex        =   18
      Top             =   1165
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   8250
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   8250
      TabIndex        =   19
      Top             =   1620
      Width           =   1100
   End
   Begin VB.Frame fra������ 
      Caption         =   "�ļ�����������"
      Height          =   2295
      Left            =   90
      TabIndex        =   21
      Top             =   180
      Width           =   8025
      Begin VB.CommandButton cmdAccessDir 
         Caption         =   "��"
         Height          =   270
         Left            =   7650
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   255
      End
      Begin MSComCtl2.UpDown upd��� 
         Height          =   300
         Left            =   7680
         TabIndex        =   9
         Top             =   660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt���������"
         BuddyDispid     =   196625
         OrigLeft        =   7695
         OrigTop         =   660
         OrigRight       =   7935
         OrigBottom      =   915
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt��������� 
         Height          =   300
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   660
         Width           =   345
      End
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   4
         Top             =   675
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   6
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txtAccessDir 
         Height          =   300
         Left            =   1200
         MaxLength       =   500
         TabIndex        =   1
         Top             =   300
         Width           =   6705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&N)"
         Height          =   350
         Left            =   150
         Picture         =   "frmFilesSet.frx":058A
         TabIndex        =   10
         Top             =   1110
         Width           =   960
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&M)"
         Height          =   350
         Left            =   150
         TabIndex        =   11
         Top             =   1455
         Width           =   960
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   150
         TabIndex        =   12
         Top             =   1800
         Width           =   960
      End
      Begin MSComctlLib.ListView lvwFileServer 
         Height          =   1095
         Left            =   1170
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1931
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils"
         SmallIcons      =   "ils"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "վ�����Ŀ¼"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "վ������û�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "վ���������"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblFileNo 
         Caption         =   "���������"
         Height          =   225
         Left            =   6435
         TabIndex        =   7
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblPassWord 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   3270
         TabIndex        =   5
         Top             =   735
         Width           =   720
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "�����û���"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   735
         Width           =   900
      End
      Begin VB.Label lblAccessDir 
         AutoSize        =   -1  'True
         Caption         =   "����Ŀ¼"
         Height          =   180
         Left            =   420
         TabIndex        =   0
         Top             =   360
         Width           =   720
      End
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   3405
      TabIndex        =   23
      Top             =   6390
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   6255
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFilesSet.frx":6DDC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "9:52"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ils 
      Left            =   8730
      Top             =   1050
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
            Picture         =   "frmFilesSet.frx":766E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFilesSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum HeadInfor
    ��� = 0
    ������
    �汾��
    �޸�����
    ��Ϣ
    ��������
    ˵��
    ����
    ��װ·��
    MD5
    �ռ�����
End Enum

Private mblnReturn As Boolean
Private mblnChangeDirectory As Boolean      '�Ƿ�ı�Ŀ¼
Private mblnAutoSet As Boolean     '�Զ���������(�����Զ��ռ��ļ����Զ����汾���������ļ��嵥���Զ������еĿͻ���Ĭ��ΪҪ����)
Private mblnFirst As Boolean
Private mblnSourceCode As Boolean '��Դ����ִ��
Private Const mstrzlAppSoftPath = "C:\AppSoft"
Private mstrSourceFloder As String '��ʱ�ռ�Ŀ¼
Public mobjFile As New FileSystemObject
Public mblnOptType As Boolean 'False �ռ����� True �ռ�����

Private Sub cmdAccessDir_Click()
    Dim strFolderName As String
    
    strFolderName = OpenFolder(Me, "ѡ�����²���������Ŀ¼")
    If strFolderName = "" Then SetCtlEnable True: Exit Sub
    If Len(strFolderName) = 3 Then
        MsgBox "����ѡ���Ŀ¼(" & strFolderName & ")!", vbInformation + vbDefaultButton1, gstrSysName
        SetCtlEnable True
        Exit Sub
    End If
    err = 0
    Me.txtAccessDir.Tag = Trim(strFolderName)
    
    If InStr(1, strFolderName, "\\") <> 0 Then
        Me.txtAccessDir.Text = strFolderName
    Else
        Me.txtAccessDir.Text = "\\" & GetMyCompterName & Mid(strFolderName, 3)
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim objItem As ListItem
    If CheckFileServer = False Then Exit Sub
    With lvwFileServer
        Set objItem = .ListItems.Add(, "K" & txt���������.Text, txt���������.Text, 1, 1)
        objItem.Selected = True
        objItem.SubItems(1) = Trim(txtAccessDir.Text)
        objItem.SubItems(2) = Trim(txtUserName)
        objItem.SubItems(3) = Trim(txtPassword)
        objItem.Tag = "1"
    End With
    Call SetFileSeverCtrlEnable
End Sub

Private Sub CmdDelete_Click()
    '����:ɾ��������
    Dim lngIndex As Long
    With lvwFileServer
        If .SelectedItem Is Nothing Then Exit Sub
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        If lngIndex >= lvwFileServer.ListItems.Count And lvwFileServer.ListItems.Count <> 0 Then
            .ListItems(.ListItems.Count).Selected = True
            .SelectedItem.EnsureVisible
        ElseIf lvwFileServer.ListItems.Count <> 0 Then
            .ListItems(lngIndex).Selected = True
            .SelectedItem.EnsureVisible
        End If
    End With
    Call SetFileSeverCtrlEnable
End Sub
Private Sub SetFileSeverCtrlEnable()
    '-------------------------------------------------------------------------------------------------
    '����:�����ļ���������ؿؼ�������ֵ
    '-------------------------------------------------------------------------------------------------
    Dim blnSel  As Boolean
    blnSel = Not Me.lvwFileServer.SelectedItem Is Nothing
    cmdModify.Enabled = blnSel
    cmdDelete.Enabled = blnSel
End Sub
Private Sub cmdModify_Click()
    Dim objItem As ListItem
    If CheckFileServer(True) = False Then Exit Sub
    With lvwFileServer
        If .SelectedItem Is Nothing Then Exit Sub
        Set objItem = .SelectedItem
        objItem.Key = "K" & txt���������.Text
        objItem.Text = txt���������.Text
        objItem.SubItems(1) = Trim(txtAccessDir.Text)
        objItem.SubItems(2) = Trim(txtUserName)
        objItem.SubItems(3) = Trim(txtPassword)
        objItem.Tag = "1"
    End With
    Call SetFileSeverCtrlEnable
End Sub
Private Function CopyFileToServer(ByVal strFileServer As String, ByVal strSourcePath As String, ByVal strSharePath As String, _
    Optional ByVal strUserName As String, Optional ByVal strPassword As String, Optional ByRef strErrInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '����:�����ļ���ָ���ķ�����
    '����:strFileServer-�ļ����������
    '     strSourcePath-Դ�ļ�Ŀ¼
    '     strSharePath-�������Ĺ���Ŀ¼
    '     strUserName-���ʵ��û���
    '     strPassWord-����
    '����:strErrInfor-���صĴ�����Ϣ
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/06
    '---------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    
    '��һ��:�ȼ����ص�Ŀ¼�Ƿ����
    '     1.���ԴĿ¼�Ƿ����
    If objFile.FolderExists(strSourcePath) = False Then
        strErrInfor = "Դ�ļ�Ŀ¼:" & strSourcePath & "������,����!"
        Exit Function
    End If
    '     2.��鹲��Ŀ¼�Ƿ����
    If objFile.FolderExists(strSharePath) = False Then
        '����ȥ����,���Ƿ����
        If IsNetServer(strSharePath, strUserName, strPassword) = False Then
            strErrInfor = "�ļ�������Ŀ¼:" & strSharePath & "������,����!"
            Exit Function
        End If
        If objFile.FolderExists(strSharePath) = False Then
            strErrInfor = "�ļ�������Ŀ¼:" & strSharePath & "������,����!"
            Exit Function
        End If
    Else
        '����Ƿ����дȨ��
        If funCanWrite(strSharePath) = False Then
            strErrInfor = "�ļ�������Ŀ¼:[" & strSharePath & "]������дȨ��!" & vbCrLf & "������д��Ȩ�޻����ֶ������ļ�!"
            Exit Function
        End If
    End If
    '   3.����Ƿ�Դ�ļ����ļ��������Ƿ�һ��
    If UCase(strSharePath) = UCase(strSourcePath) Then
        '�ò����ٽ����ļ���������
        CopyFileToServer = True
        Exit Function
    End If
    
    '�ڶ���:������ͻ��˷��ʷ������Ĺ���Ŀ¼�е���������
    If mblnOptType = False Then
        err = 0: On Error Resume Next
        objFile.DeleteFolder strSharePath & "\*", True
        objFile.DeleteFile strSharePath & "\*.*", True
    End If
    err = 0: On Error GoTo ErrHand:
    '������:������ص��ļ���ָ�����ļ�Ŀ¼��
    '      1.��������Ŀ¼:Common
    If CopyFileToTargetServer(strSourcePath, strSharePath, "���ڿ���������������Ŀ¼[" & strFileServer & "]", True, strErrInfor) = False Then
        Exit Function
    End If
    
    CopyFileToServer = True
    Exit Function
ErrHand:
    strErrInfor = "�����޿�Ԥ֪�Ĵ���,������Ϣ����:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description
End Function

Public Function CopyFileToTargetServer(ByVal strSourcePath As String, ByVal strTargetPath As String, _
    ByVal strProcessCaption As String, Optional blnChildFolder As Boolean = False, Optional ByRef strErrInfor As String) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:�����ݿ�����ָ���ķ�������ȥ
    '����:strSourcePath-Դ�ļ�·��
    '     strTargetPath-Ŀ���ļ�·��
    '     strProcessCaption-��������ʾ����
    '     blnChildFolder-�Ƿ����Դ�ļ�����Ŀ¼Ҳһ��������Ŀ��·����,true- ����,false-������
    '����:strErrInfor-������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2007/09/06
    '--------------------------------------------------------------------------------------------------------
    Dim strSourceFile As String, strTarGetFile As String, strTemp As String
    Dim objFile As New FileSystemObject
    Dim i As Long
    
    stbThis.Panels(2).Text = strProcessCaption
    pgbState.Left = stbThis.Panels(2).Left + TextWidth(strProcessCaption) + 100
    pgbState.Width = stbThis.Panels(3).Left - pgbState.Left - 100
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    pgbState.Visible = True
    '�ж�commonĿ¼�Ƿ����
    If FindFile(strTargetPath) = False Then
        '����commonĿ¼
        err = 0
        On Error Resume Next
        MkDir strTargetPath
        If err <> 0 Then
            strErrInfor = "����Ŀ¼:" & strTargetPath & "ʧ��," & vbCrLf & "����������ص�Ȩ��,����ϵͳԱ����!"
            err.Clear
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    With FileList
        .Refresh
        .Path = strSourcePath
        .FileName = "*.*"
        pgbState.Max = .ListCount + IIf(blnChildFolder, 1, 0)
        pgbState.Min = 0
        For i = 0 To .ListCount - 1
            strTemp = "\" & .List(i)
            strSourceFile = strSourcePath & strTemp
            strTarGetFile = strTargetPath & strTemp
            err = 0: On Error Resume Next
            '�ļ�����,�����ж�
            If ISCopyFile(strSourceFile, strTarGetFile) = True Then
                objFile.CopyFile strSourceFile, strTarGetFile, True
            End If
            If err <> 0 Then
                If MsgBox("Դ�ļ���" & strSourceFile & vbCrLf & " ���ܿ�����Ŀ���ļ���" & vbCrLf & strTarGetFile & vbCrLf & "��,�Ƿ������" & vbNewLine & err.Description, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    pgbState.Visible = False
                    stbThis.Panels(2).Text = ""
                    Exit Function
                End If
            End If
            pgbState.value = i + 1
            DoEvents
        Next
'        If blnChildFolder Then
'            '����Դ·���µ���Ŀ¼���ݸ���ص�Ŀ��Ŀ¼��:Դ·���п��ܲ�������Ŀ¼,���������صĴ���
'            Err = 0: On Error Resume Next
'            objFile.CopyFolder strSourcePath & "\*", strTargetPath
'            Err = 0: On Error GoTo 0
'            pgbState.Value = pgbState.Value + 1
'            DoEvents
'        End If
    End With
    pgbState.Visible = False
    CopyFileToTargetServer = True
End Function

Private Sub cmdSaveInfo_Click()
'������ص����õ����ݿ���
    If IsValid = False Then Exit Sub
    If Not SaveFile Then Exit Sub
    mblnReturn = True
    stbThis.Panels(2).Text = "������������Ϣ����ɹ�!"
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Height < 7035 Then
        Me.Height = 7035
    End If
    If Me.Width < 9615 Then
        Me.Width = 9615
    End If
    
    With cmdSaveInfo
        .Left = ScaleWidth - .Width - 100
    End With
    
    With cmdRefresh
        .Left = ScaleWidth - .Width - 100
    End With
    
    With cmdSave
        .Left = ScaleWidth - .Width - 100
        cmdCancel.Left = .Left
        cmdHelp.Left = .Left
'        cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - IIf(stbThis.Visible, stbThis.Height, 0) - 50


        chk��������.Left = .Left
        DTP��������.Left = .Left
        cmd��������.Left = .Left
        
        
        cmd��������.Top = Me.ScaleHeight - cmdHelp.Height - IIf(stbThis.Visible, stbThis.Height, 0) - 50
        DTP��������.Top = cmd��������.Top - DTP��������.Height - 50
        chk��������.Top = DTP��������.Top - chk��������.Height - 50
        
    End With
    
 
    
    
    With fra������
        .Width = cmdSave.Left - .Left - 50
        txtAccessDir.Width = .Width - txtAccessDir.Left - 50
        cmdAccessDir.Left = txtAccessDir.Left + txtAccessDir.Width - cmdAccessDir.Width
        
        upd���.Left = txtAccessDir.Left + txtAccessDir.Width - upd���.Width
        txt���������.Left = upd���.Left - txt���������.Width
        lblFileNo.Left = txt���������.Left - lblFileNo.Width
        lvwFileServer.Width = txtAccessDir.Width
    End With
    
    With fra�ļ�����
        .Width = fra������.Width
        .Left = fra������.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        mshBill.Height = .Height - mshBill.Top - 60
        mshBill.Width = .Width - mshBill.Left - 60
        mshBillShow.Left = mshBill.Left
        mshBillShow.Top = mshBill.Top
        mshBillShow.Height = mshBill.Height
        mshBillShow.Width = mshBill.Width
    End With
    
    With OptType(0)
        .Left = cmdCancel.Left
        .Top = fra�ļ�����.Top + 100
    End With
    
    With OptType(1)
        .Left = cmdCancel.Left
        .Top = OptType(0).Top + OptType(0).Height + 150
    End With
End Sub

Private Sub lvwFileServer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    upd���.value = Val(Item.Text)
    txtAccessDir.Text = Item.SubItems(1)
    txtUserName.Text = Item.SubItems(2)
    txtPassword.Text = Item.SubItems(3)
    Call SetFileSeverCtrlEnable
End Sub

Private Sub cmdCancel_Click()
'    If mobjFile.FolderExists(mstrSourceFloder) Then
'        mobjFile.DeleteFolder mstrSourceFloder, True
'    End If
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdRefresh_Click()
    '�ռ��ļ���Ŀ¼
    Dim strSourceFile As String
    SetCtlEnable False
    cmdSave.Enabled = False
    mshBill.Visible = True
    mshBillShow.Visible = False
    OptType(0).Enabled = False
    OptType(1).Enabled = False
    If GetFileInforamtion() = True Then
        '���7z.exe����ϵͳ����
        Call fun_KillProcess(PROAPPCTION)
        '�����ռ��ļ���������
        Call FloderToClipBoard(mstrSourceFloder)
        '�ռ�����
        Call BillFileSort
        
        Me.cmdSave.Enabled = True
        
        strSourceFile = mstrSourceFloder & "\"
        If mobjFile.FolderExists(strSourceFile) Then
            With FileList
                .Refresh
                .Path = strSourceFile
                .FileName = "*.*"
                
                If .ListCount = 0 Then
                    MsgBox "û���ռ����κ��ļ�," & vbCrLf & "�����ļ�MD5���������һ��!", vbInformation + vbDefaultButton1 + vbOKOnly
                    cmdSave.Enabled = False
                Else
                    MsgBox "�ļ��ռ��ɹ�,��ʱ�ռ��ļ��ѿ�������������," & vbCrLf & "������������Ŀ¼û��дȨ��ʱ,��ֱ��ճ���ռ��ļ�." & vbCrLf & "ע��:�ϴ���ɺ͹رն���ɾ����ʱ�ռ�Ŀ¼�ͼ�����!", vbInformation + vbDefaultButton1 + vbOKOnly
                    cmdSave.Enabled = True
                End If
            End With
        End If
            
        mblnChangeDirectory = True
    End If
    OptType(0).Enabled = True
    OptType(1).Enabled = True
    SetCtlEnable True
End Sub

Private Sub cmdSave_Click()
    Dim strErrMsg As String
    Dim objItem As ListItem
    Dim strSourcePath As String
    '1.������ݵĺϷ���
    If IsValid = False Then Exit Sub

    '2.��Ҫ����ص��ļ��ֲ�����صķ�������
    strSourcePath = Trim(mstrSourceFloder)
    For Each objItem In lvwFileServer.ListItems
        If CopyFileToServer(objItem.Text, strSourcePath, objItem.SubItems(1), objItem.SubItems(2), objItem.SubItems(3), strErrMsg) = False Then
            MsgBox strErrMsg, vbDefaultButton1 + vbInformation, gstrSysName
            stbThis.Panels(2).Text = strErrMsg
            Exit Sub
        End If
    Next
    
    '3.������ص����õ����ݿ���
    If Not SaveFile Then Exit Sub
    mobjFile.DeleteFolder strSourcePath, True
    mblnReturn = True
    Unload Me
End Sub

Private Function IsValid() As Boolean
    '--------------------------------------------------------------------
    '����:��֤���ݵĺϷ���
    '--------------------------------------------------------------------
    Dim objItem As ListItem
    
    IsValid = False
    If mblnChangeDirectory = True Then
        If FindFile(Trim(mstrSourceFloder)) = False Then
            MsgBox "�����ļ���ָ��Ŀ¼������,����������!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    
    If InStr(1, mstrSourceFloder, "'") <> 0 Then
        MsgBox "�����ļ��в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    If lvwFileServer.ListItems.Count = 0 Then
        MsgBox "û��������ص��ļ�������,��������վ�������ķ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtAccessDir.Enabled Then txtAccessDir.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function CheckFileServer(Optional blnModify As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------
    '����:��鵱ǰ���ļ��������Ƿ���ȷ
    '����:��ǰ���ļ��������ĸ�����ȷ,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/06
    '-----------------------------------------------------------------------------
    Dim objItem As ListItem
    
    err = 0: On Error GoTo ErrHand:
    CheckFileServer = False
    If Trim(txtAccessDir.Text) = "" Then
        MsgBox "δ����վ����ʵķ�����Ŀ¼,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtAccessDir.Enabled Then txtAccessDir.SetFocus
        Exit Function
    End If
    If Trim(txtUserName.Text) = "" Then
        MsgBox "�����û�δ����,�����÷������û���!", vbInformation + vbDefaultButton1, gstrSysName
        If txtUserName.Enabled Then txtUserName.SetFocus
        Exit Function
    End If
    If InStr(1, txtUserName.Text, "'") <> 0 Then
        MsgBox "�����û��в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtUserName.Enabled Then txtUserName.SetFocus
        Exit Function
    End If
    If InStr(1, txtPassword.Text, "'") <> 0 Then
        MsgBox "���������в��ܴ��ڵ�����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtPassword.Enabled Then txtPassword.SetFocus
        Exit Function
    End If
    
    If Right(Trim(txtAccessDir.Text), 1) = "\" Then
        txtAccessDir.Text = Left(txtAccessDir.Text, Len(txtAccessDir.Text) - 1)
    End If
    
    For Each objItem In lvwFileServer.ListItems
        If blnModify = True Then
            If Val(objItem.Text) = Val(txt���������.Text) And objItem.Selected = False Then
                MsgBox "���������Ϊ" & txt���������.Text & "�Ѿ�����,���������ô˱�ŵķ�����!", vbInformation + vbDefaultButton1, gstrSysName
                If txt���������.Enabled Then txt���������.SetFocus
                Exit Function
            End If
            If objItem.SubItems(1) = txtAccessDir.Text And objItem.Selected = False Then
                MsgBox "������ͬ�ķ���Ŀ¼,�ò���������!", vbInformation + vbDefaultButton1, gstrSysName
                If txtAccessDir.Enabled Then txtAccessDir.SetFocus
                Exit Function
            End If
        Else
            If Val(objItem.Text) = Val(txt���������.Text) Then
                MsgBox "���������Ϊ" & txt���������.Text & "�Ѿ�����,���������ô˱�ŵķ�����!", vbInformation + vbDefaultButton1, gstrSysName
                If txt���������.Enabled Then txt���������.SetFocus
                Exit Function
            End If
            If objItem.SubItems(1) = txtAccessDir.Text Then
                MsgBox "������ͬ�ķ���Ŀ¼,�ò���������!", vbInformation + vbDefaultButton1, gstrSysName
                If txtAccessDir.Enabled Then txtAccessDir.SetFocus
                Exit Function
            End If
        End If
    Next
    
    If FindFile(Trim(txtAccessDir.Text)) = False Then
        If IsNetServer(Trim(txtAccessDir.Text), Trim(txtUserName.Text), Trim(txtPassword.Text)) = False Then
            MsgBox "�����ļ���ָ��Ŀ¼������,����������!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        Call CancelNetServer(Trim(txtAccessDir.Text))
    End If
    CheckFileServer = True
    Exit Function
ErrHand:
End Function

Private Function SaveFile() As Boolean
    '-----------------------------------------------------------------------------
    '����:����ص����ñ��浽���ݿ���
    '����:����ɹ�,����true,���򷵻�False
    '����:ף��
    '����:2010/12/13
    '-----------------------------------------------------------------------------
    Dim strSQL As String, objItem As ListItem

    
    err = 0
    On Error GoTo ErrHand:
    SaveFile = False
    gcnOracle.BeginTrans

    '�������ص�����
    strSQL = "Delete zlregInfo where (��Ŀ like '������Ŀ¼%' or ��Ŀ like '�����û�%' or ��Ŀ like '��������%') and ��Ŀ not in ('�����û�S','��������S','�����û�F','��������F') "
    gcnOracle.Execute strSQL
'    strSQL = "Update zlreginfo set ����=NULL where ��Ŀ in ('������Ŀ¼','�����û�','��������')"
'    gcnOracle.Execute strSQL
    '�����µķ�������
    For Each objItem In lvwFileServer.ListItems
'        If Val(objItem.Text) = 0 Then
'            strSQL = "Update zlreginfo set ����='" & Trim(objItem.SubItems(1)) & "' where ��Ŀ='������Ŀ¼'"
'            gcnOracle.Execute strSQL
'            strSQL = "Update zlreginfo set ����='" & Trim(objItem.SubItems(2)) & "' where ��Ŀ='�����û�'"
'            gcnOracle.Execute strSQL
'            strSQL = "Update zlreginfo set ����='" & Trim(objItem.SubItems(3)) & "' where ��Ŀ='��������'"
'            gcnOracle.Execute strSQL
'        Else
            strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('������Ŀ¼" & objItem.Text & "',Null,'" & Trim(objItem.SubItems(1)) & "')"
            gcnOracle.Execute strSQL
            strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�����û�" & objItem.Text & "',Null,'" & Trim(objItem.SubItems(2)) & "')"
            gcnOracle.Execute strSQL
            strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('��������" & objItem.Text & "',Null,'" & Trim(objItem.SubItems(3)) & "')"
            gcnOracle.Execute strSQL
'        End If
    Next
    gcnOracle.CommitTrans
    SaveFile = True
    Exit Function
ErrHand:
    MsgBox "�������·�������Ϣʱ���ִ���,���ܴ���������ͬ�ķ�����!" & vbNewLine & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    gcnOracle.RollbackTrans
End Function

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    mstrSourceFloder = GetTmpPath & "TEMPGATHER"
    If Load��������Ϣ() = False Then Unload Me: Exit Sub
    
    mblnSourceCode = IsSourceCode
    Me.cmdSave.Enabled = False
    Me.mshBill.Visible = True
    Me.mshBillShow.Visible = False
    
    '����ͷ��Ϣ
    Call LoadHeadInfor
    '��ʼ������Ϣ.
    Call intBillInfor
    '�Ƚ��ռ���Ϣ
    Call CompareFile
    '�ж���������
    Call OpinionUpGradeDate
    mshBill.AutoRefresh = False

    If mblnOptType Then
        OptType(1).value = True
    Else
        OptType(0).value = True
    End If
    
    mblnChangeDirectory = False
    mblnReturn = False
    
    '�޸�Ϊ�򿪴���ʱɾ����ʱĿ¼
    If mobjFile.FolderExists(mstrSourceFloder) Then
        mobjFile.DeleteFolder mstrSourceFloder, True
    End If
    
    If mblnAutoSet Then
        '�Զ���������(�����Զ��ռ��ļ����Զ����汾���������ļ��嵥���Զ������еĿͻ���Ĭ��ΪҪ����)
        If AutoSet = False Then Exit Sub
        '�����е�վ���Ϊ����
        Call ExecuteProcedure("Zl_Zlclients_Control(4,Null,Null,1)", Me.Caption)
        Call cmdSave_Click
    End If
    Call SetFileSeverCtrlEnable
End Sub

Private Function AutoSet() As Boolean
    '------------------------------------------------------------------------------------------------------------
    '����:�Զ�����
    '����:���óɹ�,����true,���򷵻�False
    '------------------------------------------------------------------------------------------------------------
    SetCtlEnable False

    If GetFileInforamtion = False Then
        SetCtlEnable True: Exit Function
    End If
    
    SetCtlEnable True
    AutoSet = True
End Function

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    mblnFirst = True
End Sub

Private Sub LoadHeadInfor()
    '------------------------------------------------------------------------------------------------------------
    '����:����ͷ��Ϣ
    '------------------------------------------------------------------------------------------------------------
    With mshBill
        .Active = True
        .Cols = 11
        .Clear
        .Rows = 2
        '.MsfObj.FixedCols = 1
        .TextMatrix(0, HeadInfor.���) = "���"
        .TextMatrix(0, HeadInfor.������) = "������"
        .TextMatrix(0, HeadInfor.�汾��) = "�汾��"
        .TextMatrix(0, HeadInfor.�޸�����) = "�޸�����"
        .TextMatrix(0, HeadInfor.��Ϣ) = "��Ϣ"
        .TextMatrix(0, HeadInfor.��������) = "��������"
        .TextMatrix(0, HeadInfor.˵��) = "˵��"
        .TextMatrix(0, HeadInfor.����) = "����"
        .TextMatrix(0, HeadInfor.��װ·��) = "��װ·��"
        .TextMatrix(0, HeadInfor.�ռ�����) = "�ռ�����"
        
        .ColWidth(HeadInfor.���) = 500
        .ColWidth(HeadInfor.������) = 2000
        .ColWidth(HeadInfor.�汾��) = 900
        .ColWidth(HeadInfor.�޸�����) = 1800
        .ColWidth(HeadInfor.��Ϣ) = 2000
        .ColWidth(HeadInfor.��������) = 1800
        .ColWidth(HeadInfor.˵��) = 2000
        .ColWidth(HeadInfor.����) = 800
        .ColWidth(HeadInfor.��װ·��) = 2000
        .ColWidth(HeadInfor.MD5) = 0
        .ColWidth(HeadInfor.�ռ�����) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        
        .ColData(HeadInfor.���) = 5
        .ColData(HeadInfor.������) = 5
        .ColData(HeadInfor.�汾��) = 5
        .ColData(HeadInfor.�޸�����) = 5
        .ColData(HeadInfor.��Ϣ) = 5
        .ColData(HeadInfor.��������) = 5
        .ColData(HeadInfor.˵��) = 5
        .ColData(HeadInfor.����) = 5
        .ColData(HeadInfor.��װ·��) = 5
        .ColData(HeadInfor.MD5) = 5
        .ColData(HeadInfor.�ռ�����) = 5
        
        .ColAlignment(HeadInfor.���) = flexAlignCenterCenter
        .ColAlignment(HeadInfor.������) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.�汾��) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.�޸�����) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.��Ϣ) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.��������) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.˵��) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.����) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.��װ·��) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.MD5) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.�ռ�����) = flexAlignLeftCenter
        
        .Active = False
    End With
End Sub

Private Sub LoadHeadInforShow()
    '------------------------------------------------------------------------------------------------------------
    '����:����ͷ��Ϣ
    '------------------------------------------------------------------------------------------------------------
    With mshBillShow
        .Active = True
        .Cols = 11
        .Clear
        .Rows = 2
        '.MsfObj.FixedCols = 1
        .TextMatrix(0, HeadInfor.���) = "���"
        .TextMatrix(0, HeadInfor.������) = "������"
        .TextMatrix(0, HeadInfor.�汾��) = "�汾��"
        .TextMatrix(0, HeadInfor.�޸�����) = "�޸�����"
        .TextMatrix(0, HeadInfor.��Ϣ) = "��Ϣ"
        .TextMatrix(0, HeadInfor.��������) = "��������"
        .TextMatrix(0, HeadInfor.˵��) = "˵��"
        .TextMatrix(0, HeadInfor.����) = "����"
        .TextMatrix(0, HeadInfor.��װ·��) = "��װ·��"
        .TextMatrix(0, HeadInfor.�ռ�����) = "�ռ�����"
        
        .ColWidth(HeadInfor.���) = 500
        .ColWidth(HeadInfor.������) = 2000
        .ColWidth(HeadInfor.�汾��) = 900
        .ColWidth(HeadInfor.�޸�����) = 1800
        .ColWidth(HeadInfor.��Ϣ) = 2000
        .ColWidth(HeadInfor.��������) = 1800
        .ColWidth(HeadInfor.˵��) = 2000
        .ColWidth(HeadInfor.����) = 800
        .ColWidth(HeadInfor.��װ·��) = 2000
        .ColWidth(HeadInfor.MD5) = 0
        .ColWidth(HeadInfor.�ռ�����) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        
        .ColData(HeadInfor.���) = 5
        .ColData(HeadInfor.������) = 5
        .ColData(HeadInfor.�汾��) = 5
        .ColData(HeadInfor.�޸�����) = 5
        .ColData(HeadInfor.��Ϣ) = 5
        .ColData(HeadInfor.��������) = 5
        .ColData(HeadInfor.˵��) = 5
        .ColData(HeadInfor.����) = 5
        .ColData(HeadInfor.��װ·��) = 5
        .ColData(HeadInfor.MD5) = 5
        .ColData(HeadInfor.�ռ�����) = 5
        
        .ColAlignment(HeadInfor.���) = flexAlignCenterCenter
        .ColAlignment(HeadInfor.������) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.�汾��) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.�޸�����) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.��Ϣ) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.��������) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.˵��) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.����) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.��װ·��) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.MD5) = flexAlignLeftCenter
        .ColAlignment(HeadInfor.�ռ�����) = flexAlignLeftCenter
        
        .Active = False
    End With
End Sub

Private Function GetFileInforamtion() As Boolean
        '------------------------------------------------------------------------
        '--����:��ȡ���²�����Ϣ
        '--����:���سɹ�,����true,����false
        '------------------------------------------------------------------------
        Dim strCurFileDirectory As String
        Dim lngRow As Long
        Dim lngErr As Long
        
'        Dim strSQL As String
'        Dim rsTmp As New ADODB.Recordset
        Dim strPath As String '����װ·��
        Dim strFullPath As String '��װ·��
        
        Dim strCompTxt  As String 'ѹ���ű�
        Dim strSource   As String 'ѹ��Դ�ļ�
        Dim strDesc     As String 'ѹ��Ŀ���ļ�
'        Dim RetVal      As Long  '����ֵ
        Dim objFile As New FileSystemObject
        Dim usrUpList  As UpdateList
        Dim lngSuccess  As Long
        Dim str7zFile   As String
        
        '���ݿ��ļ���ֵ
        Dim strFilename As String '�ļ���
        Dim strFileType As String '�ļ�����
        Dim strSetupPath As String '��װ·��
        Dim strFileMD5   As String '�ļ�MD5ֵ
        
        Dim strLocaFileMD5 As String '�����ļ�MD5ֵ
        Dim driver As Drive

        err = 0: On Error GoTo ErrHand:
        strCurFileDirectory = Trim(mstrSourceFloder)
        GetFileInforamtion = False
        
        '���ʣ��ռ�
        For Each driver In objFile.Drives
            If driver.IsReady Then
                If driver.DriveLetter = "C" Then
                    If driver.FreeSpace < 204800000 Then 'С��200M
                        MsgBox "��ʱ�ռ�Ŀ¼û���㹻�Ŀռ�!", vbInformation, gstrSysName
                        Exit Function
                    End If
                    Exit For
                End If
            End If
        Next driver

        If FindFile(strCurFileDirectory) = False Then
            On Error Resume Next
            Call mobjFile.CreateFolder(strCurFileDirectory)
            If mobjFile.FolderExists(strCurFileDirectory) = False Then
                MsgBox "��ʱ�ռ�Ŀ¼���ܴ���,����!" & vbCrLf & strCurFileDirectory, vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        End If
        
        '2��ѹ���ļ�
        str7zFile = GetWinSystemPath & "\7z.exe"
        If FindFile(str7zFile) = False Then
            MsgBox "ѹ���ļ�7z.exe������,���ֶ�����ϵͳĿ¼��!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        
        str7zFile = GetWinSystemPath & "\7z.dll"
        If FindFile(str7zFile) = False Then
            MsgBox "ѹ���ļ�7z.dll������,���ֶ�����ϵͳĿ¼��!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
       
       
        '�������ʱ�ռ��ļ�Ŀ¼�е���������
        err = 0: On Error Resume Next
        objFile.DeleteFolder strCurFileDirectory & "\*", True
        objFile.DeleteFile strCurFileDirectory & "\*.*", True
        
        If mblnSourceCode Then
            strPath = mstrzlAppSoftPath
        Else
            strPath = App.Path
        End If
        
        
        pgbState.Visible = True
        stbThis.Panels(2).Text = "�����ռ���ѹ������"
        pgbState.Left = stbThis.Panels(2).Left + TextWidth("�����ռ���ѹ������") + 100
        pgbState.Width = stbThis.Panels(3).Left - pgbState.Left - 100
        pgbState.Top = stbThis.Top + stbThis.Height / 3

        With mshBill
        If .Rows = 0 Then Exit Function
        '        DoEvents
        pgbState.Max = .Rows - 1
        pgbState.Min = 0
        pgbState.value = 0
        
        Erase usrUpList.uFile
        lngSuccess = 0
        lngErr = 0
        
        For lngRow = 1 To .Rows - 1
                strFilename = .TextMatrix(lngRow, HeadInfor.������)
                strFileType = .TextMatrix(lngRow, HeadInfor.����)
                strSetupPath = .TextMatrix(lngRow, HeadInfor.��װ·��)
                strFileMD5 = .TextMatrix(lngRow, HeadInfor.MD5)
'                If strFileName = UCase("zl9MediStore.dll") Then
'                    MsgBox 1
'                End If
                '��ȡ������·��
                strFullPath = GetSetupPath(Nvl(strFilename, ""), Nvl(strSetupPath, ""), Nvl(strFileType, ""), strPath)
                If strFullPath = "" Then
                    If Nvl(strFilename, "") <> "" Then
                        .TextMatrix(lngRow, HeadInfor.��Ϣ) = "δָ��·��!"
                        .TextMatrix(lngRow, HeadInfor.�ռ�����) = "1"
                        .SetRowColor lngRow, vbRed, False
                        lngErr = lngErr + 1
                    End If
                Else
                    '7z����ѹ��
                    '4���ļ�����Ҫѹ�� ���⴦��
                    If UCase(Nvl(strFilename, "")) = UCase("zlHisCrust.exe") Or UCase(Nvl(strFilename, "")) = UCase("7z.exe") Or UCase(Nvl(strFilename, "")) = UCase("7z.dll") Or UCase(Nvl(strFilename, "")) = UCase("aamd532.dll") Or UCase(Nvl(strFilename, "")) = UCase("zlRunas.exe") Or UCase(Nvl(strFilename, "")) = UCase("RegCom.dll") Then
                        strDesc = strCurFileDirectory & "\" & strFilename
                        If FindFile(strFullPath) Then
                            strLocaFileMD5 = HashFile(strFullPath, 2 ^ 27)  '��¼MD5�Ա�ͻ�������ʱ�����ļ��Ƚ�
                            If OptType(1).value Then
                                If UCase(strFileMD5) = UCase(strLocaFileMD5) Then
                                  GoTo Success
                                End If
                            End If
                            ReDim Preserve usrUpList.uFile(lngSuccess)
                            usrUpList.uFile(lngSuccess).FileName = strFilename
                            usrUpList.uFile(lngSuccess).FileVision = GetCommpentVersion(strFullPath)
                            usrUpList.uFile(lngSuccess).FileEditDate = Format(FileDateTime(strFullPath), "yyyy-MM-DD hh:mm:ss")
                            usrUpList.uFile(lngSuccess).FileMD5 = strLocaFileMD5
                            
                            If ISCopyFile(strFullPath, strDesc) = True Then
                                objFile.CopyFile strFullPath, strDesc, True
                            End If
Success:
                            .TextMatrix(lngRow, HeadInfor.��Ϣ) = ""
                            .TextMatrix(lngRow, HeadInfor.�ռ�����) = "0"
                            .SetRowColor lngRow, &HFFFFFF, False
                            lngSuccess = lngSuccess + 1
                        Else
                             .TextMatrix(lngRow, HeadInfor.��Ϣ) = "δ��װ�ļ�!"
                             .TextMatrix(lngRow, HeadInfor.�ռ�����) = "2"
                             .SetRowColor lngRow, vbRed, False
                             lngErr = lngErr + 1
                        End If
                    Else
                        If FindFile(strFullPath) Then
                            strLocaFileMD5 = HashFile(strFullPath, 2 ^ 27)  '��¼MD5�Ա�ͻ�������ʱ�����ļ��Ƚ�
                            If OptType(1).value Then
                                If UCase(strFileMD5) = UCase(strLocaFileMD5) Then
                                  GoTo Success1
                                End If
                            End If
                            
                            ReDim Preserve usrUpList.uFile(lngSuccess)
                            usrUpList.uFile(lngSuccess).FileName = strFilename
                            usrUpList.uFile(lngSuccess).FileVision = GetCommpentVersion(strFullPath)
                            usrUpList.uFile(lngSuccess).FileEditDate = Format(FileDateTime(strFullPath), "yyyy-MM-DD hh:mm:ss")
                            usrUpList.uFile(lngSuccess).FileMD5 = strLocaFileMD5
                            
                           
                            strSource = strFullPath
                            strDesc = strCurFileDirectory & "\" & GetCompressName(Nvl(strFilename, ""))
                            strCompTxt = CompressionCmd(strDesc, strSource, COMPRESSIONRATE)
                            If strCompTxt <> "" Then
'                                RetVal = Shell(strCompTxt, vbHide)
                                 Call GetCmdTxt(strCompTxt)
                            End If
Success1:
                            .TextMatrix(lngRow, HeadInfor.��Ϣ) = ""
                            .TextMatrix(lngRow, HeadInfor.�ռ�����) = "0"
                            .SetRowColor lngRow, &HFFFFFF, False
                            lngSuccess = lngSuccess + 1
                        Else
                             .TextMatrix(lngRow, HeadInfor.��Ϣ) = "δ��װ�ļ�!"
                             .TextMatrix(lngRow, HeadInfor.�ռ�����) = "2"
                             .SetRowColor lngRow, vbRed, False
                             lngErr = lngErr + 1
                        End If
                    End If
                End If

'                .ListCount = lngRow
                DoEvents
                If pgbState.value >= pgbState.Max Then
                    pgbState.value = pgbState.Max
                Else
                    pgbState.value = pgbState.value + 1
                End If
            Next
        End With
        
        '���������ű�
        Call SaveUpList(usrUpList)
        
        pgbState.Visible = False
        If lngErr = 0 Then
            stbThis.Panels(2).Text = ""
        Else
            stbThis.Panels(2).Text = lngErr & "���ļ�δ��װ"
        End If
        GetFileInforamtion = True
        
        Exit Function
ErrHand:
        MsgBox "���ռ��ļ�ʱ,�����˴���,������Ŀ���ļ�" & vbCrLf & "�Ѿ�������,������ϢΪ:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
        pgbState.Visible = False
        stbThis.Panels(2).Text = ""
        GetFileInforamtion = False
        
End Function

Private Function FindFile(ByVal strFilename As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--����:����ָ�����ļ����ļ��Ƿ����
    '--����: ������ڴ��ļ�ΪTrue,����ΪFlase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFilename) > 0 Then
        apiOpenFile strFilename, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Private Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--����:��ȡϵͳĿ¼
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Dim StrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    StrWinPath = Left(Buffer, rtn)
    GetWinPath = StrWinPath
End Function
Private Function GetWinSystemPath() As String
    
    Dim Buffer As String
    Dim strSystem As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetSystemDirectory(Buffer, Len(Buffer))
    strSystem = Left(Buffer, rtn)
    
    GetWinSystemPath = strSystem
End Function
Private Function Load��������Ϣ() As Boolean
    '---------------------------------------------------------------------------------------------------
    '����:���ط�������Ϣ
    '����:
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/06
    '---------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem
    Dim str�������� As String
    
    
    err = 0: On Error GoTo ErrHand:
    Set rsTmp = New ADODB.Recordset
    
    gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ like '������Ŀ¼%'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Me.lvwFileServer.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            str�������� = Replace(Nvl(rsTemp!��Ŀ), "������Ŀ¼", "")
                If str�������� <> "" Then
                Set objItem = lvwFileServer.ListItems.Add(, "K" & Val(str��������), Val(str��������), 1, 1)
                objItem.SubItems(1) = Nvl(rsTemp!����)
                
                gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ ='�����û�" & str�������� & "'"
                Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
                If rsTmp.EOF = False Then
                    objItem.SubItems(2) = Nvl(rsTmp!����)
                Else
                    objItem.SubItems(2) = ""
                End If
                
                gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ ='��������" & str�������� & "'"
                Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
                If rsTmp.EOF = False Then
                    objItem.SubItems(3) = Nvl(rsTmp!����)
                Else
                    objItem.SubItems(3) = ""
                End If
                objItem.Tag = ""
            End If
            .MoveNext
        Loop
        .Close
    End With
    If Not lvwFileServer.SelectedItem Is Nothing Then
        lvwFileServer.SelectedItem.Selected = False
        Set lvwFileServer.SelectedItem = Nothing
    End If
    Load��������Ϣ = True
    Exit Function
ErrHand:
End Function

Public Sub ShowEdit(ByVal frmMain As Object, ByRef blnRetun As Boolean, Optional blnAutoSet As Boolean)
    '-------------------------------------------------------------------------------
    '--���ܣ���ʾ�ͱ༭������Ϣ
    '--������frmMain-������
    '       blnAutoSet-�Զ���������(�����Զ��ռ��ļ����Զ����汾���������ļ��嵥���Զ������еĿͻ���Ĭ��ΪҪ����)
    '--���أ�blnRetun-�༭�ɹ�����true,���򷵻�false
    '--      strSourceDirectory-����ָ����Դ�ļ�Ŀ¼
    '-------------------------------------------------------------------------------
    mblnAutoSet = blnAutoSet
    Me.cmdSave.Enabled = False
    
    Me.Show 1, frmMDIMain
    blnRetun = mblnReturn
End Sub
 

Private Sub lvwFileServer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mshBill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Integer
    err = 0: On Error Resume Next
    If KeyCode = vbKeyDelete Then
        If mshBill.Rows <> 2 Then
            mshBill.MsfObj.RowPosition(mshBill.MsfObj.Row) = mshBill.MsfObj.Rows - 1
            mshBill.Rows = mshBill.Rows - 1
        Else
            mshBill.ClearBill
        End If
        With mshBill
            .Redraw = False
            For i = 1 To .Rows - 1
                If .TextMatrix(i, HeadInfor.������) <> "" Then
                    .TextMatrix(i, HeadInfor.���) = i
                End If
            Next
            .Redraw = True
        End With
        cmdSave.Enabled = True
        mblnChangeDirectory = True
    End If
    
End Sub

Private Sub mshBill_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptType_Click(Index As Integer)
    If OptType(0).value Then
        mblnOptType = False
    Else
        mblnOptType = True
    End If
End Sub

Private Sub txtAccessDir_Change()
    Me.cmdSave.Enabled = True
End Sub

Private Sub txtAccessDir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtFileSource_Change()
    mblnChangeDirectory = True
End Sub

Private Sub txtFileSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtPassword_Change()
    Me.cmdSave.Enabled = True
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtUserName_Change()
    Me.cmdSave.Enabled = True
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
Private Sub intBillInfor()
    '--------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    Dim str�汾�� As String
    Dim lng�汾�� As Long
    Dim str�������� As String
    
    err = 0
    On Error Resume Next
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Zlfilesupgrade")
    With rsTmp
        lngRow = 1
        Do While Not .EOF
            mshBill.TextMatrix(lngRow, HeadInfor.���) = lngRow
            mshBill.TextMatrix(lngRow, HeadInfor.������) = IIf(IsNull(!�ļ���), "", !�ļ���)
            str�汾�� = ""
            If !�汾�� > 0 Then
                lng�汾�� = !�汾��
                str�汾�� = Int(lng�汾�� / 10 ^ 8)
                lng�汾�� = lng�汾�� Mod 10 ^ 8
                str�汾�� = str�汾�� & "." & Int(lng�汾�� / 10 ^ 4)
                lng�汾�� = lng�汾�� Mod 10 ^ 4
                str�汾�� = str�汾�� & "." & lng�汾��
            End If
            
            str�������� = IIf(IsNull(!��������), "", !��������)
            If str�������� <> "" Then
                str�������� = Format(str��������, "yyyy-MM-dd hh:mm:ss")
            End If
            
            mshBill.TextMatrix(lngRow, HeadInfor.�汾��) = str�汾��
            mshBill.TextMatrix(lngRow, HeadInfor.�޸�����) = Format(!�޸�����, "yyyy-MM-dd hh:mm:ss")
            mshBill.TextMatrix(lngRow, HeadInfor.��������) = str��������
            mshBill.TextMatrix(lngRow, HeadInfor.˵��) = IIf(IsNull(!˵��), "", !˵��)
            mshBill.TextMatrix(lngRow, HeadInfor.����) = IIf(IsNull(!����), "", !����)
            mshBill.TextMatrix(lngRow, HeadInfor.��װ·��) = IIf(IsNull(!��װ·��), "", !��װ·��)
            If IIf(IsNull(!MD5), "", !MD5) <> "" Then
                mblnOptType = True
            End If
            mshBill.TextMatrix(lngRow, HeadInfor.MD5) = IIf(IsNull(!MD5), "", !MD5)
            mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����) = "0"
            
            mshBill.Rows = mshBill.Rows + 1
            lngRow = lngRow + 1
            .MoveNext
        Loop
        If .RecordCount <> 0 Then
            mshBill.Rows = mshBill.Rows - 1
        End If
    End With
End Sub
Private Sub SetCtlEnable(Optional blnEnable As Boolean = False)
    '--------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable����
    '--------------------------------------------------------------------------------------------
    Me.cmdCancel.Enabled = blnEnable
    Me.cmdHelp.Enabled = blnEnable
    Me.txtPassword.Enabled = blnEnable
    Me.txtUserName.Enabled = blnEnable
    Me.mshBill.Enabled = blnEnable
End Sub

Private Function IsSourceCode() As Boolean
    '-----------------------------------------------------------------------------------------
    '����:ȷ���Ƿ�Դ����
    '����:��ԭ����-true,����Դ����-false
    '-----------------------------------------------------------------------------------------
    err = 0: On Error Resume Next
    Debug.Print 1 / 0
    IsSourceCode = err <> 0
End Function
Public Function ISCopyFile(ByVal strSourceFile As String, ByVal strTarGetFile As String) As Boolean
     '---------------------------------------------------------------------------------------------------------------
    '
    '����:�ж��Ƿ���Ҫ�����ļ�(�Ƚϰ汾��,�޸�ʱ��)
    '�����:
    '   strSourceFile:Դ�ļ�
    '   strTargetFile:Ŀ���ļ�
    '����:��Ҫ�����򷵻�true,���򷵻�false
    '---------------------------------------------------------------------------------------------------------------
    Dim strSource As String, strTarget As String
    
    ISCopyFile = False
    err = 0: On Error Resume Next
    If FindFile(strTarGetFile) = False Then
        'û�з����ļ����򷵻�true
        ISCopyFile = True
        Exit Function
    End If
    
    '�Ƚ��ļ��汾��
    strTarget = GetCommpentVersion(strTarGetFile)
    strSource = GetCommpentVersion(strSourceFile)
    If RtnVerNum(strTarget) < RtnVerNum(strSource) Then
        ISCopyFile = True
        Exit Function
    End If
    
    '�Ƚ��ļ�������޸�ʱ��
    strTarget = Format(FileDateTime(strTarGetFile), "yyyy-MM-DD hh:mm:ss")
    strSource = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
    If strTarget < strSource Then
        ISCopyFile = True
        Exit Function
    End If
End Function
Private Function RtnVerNum(ByVal strVer As String) As Long
    '--------------------------------------------------------------------------------------------------------------------------------
    '--����:�������ְ汾
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim strArr
    
    If strVer <> "" Then
        strArr = Split(strVer, ".")
        RtnVerNum = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
    Else
        RtnVerNum = 0
    End If
End Function
 
Private Sub txt���������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If

End Sub
Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ؼ��İ汾��
    '���:
    '����:
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:���˺�
    '����:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '��ȡ�ļ��汾��
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Private Function GetSetupPath(ByVal strFilename As String, ByVal strPathSign As String, ByVal strFileType As String, ByVal strPath As String) As String
    '--------------------------------------------------------------------------------------------------------
    '����:��ȡ�ռ��ļ�������·��
    '����:����������·��
    '����:ף��
    '����:2010/12/10
    '--------------------------------------------------------------------------------------------------------
    On Error GoTo ErrH
    Dim strTemp As String '��ʱ·�����
    Dim strSystemDirectory As String 'ϵͳsystem32Ŀ¼
    Dim strWinDirectory As String  'windowsĿ¼
    strSystemDirectory = GetWinSystemPath
    strWinDirectory = GetWinPath
    
    If strFilename = "" Then
        GetSetupPath = ""
        Exit Function
    End If
    
    If Len(strPathSign) = 0 Then
        Select Case strFileType
        Case "0" '����
            strTemp = strPath & "\Public\" & strFilename
        Case "1" 'Ӧ��
            strTemp = strPath & "\Apply\" & strFilename
        Case "2" '����
            strTemp = strWinDirectory & "\Help\" & strFilename
        Case "3" '����
            strTemp = strPath & "\" & strFilename
        Case "4" '����
            strTemp = ""
        Case "5" 'ϵͳ
            strPathSign = UCase(strPathSign)
            If (InStrRev(strPathSign, "[SYSTEM]", -1) > 0) Or (strPathSign = "") Then
                strTemp = strSystemDirectory & "\" & strFilename
            End If
            
            '��·��
            If InStrRev(strPathSign, "[PUBLIC]", -1) > 0 Then
                strTemp = strPath & "\PUBLIC\" & strFilename
            End If
        End Select
    Else
        strPathSign = UCase(strPathSign)
        If InStrRev(strPathSign, "[APPSOFT]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[APPSOFT]", strPath)
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        ElseIf InStrRev(strPathSign, "[SYSTEM]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[SYSTEM]", strSystemDirectory)
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        ElseIf InStrRev(strPathSign, "[PUBLIC]", -1) > 0 Then
            strTemp = Replace(strPathSign, "[PUBLIC]", strPath & "\PUBLIC")
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        ElseIf InStrRev(strPathSign, "[HELP]", -1) Then
            strTemp = Replace(strPathSign, "[HELP]", strWinDirectory & "\Help")
            If Right(strTemp, 1) <> "\" Then
                strTemp = strTemp & "\" & strFilename
            Else
                strTemp = strTemp & strFilename
            End If
        Else '����·��
            If Left(strFilename, 2) = "\\" Then
                strTemp = ""
            Else
                strTemp = Left(strPath, 1) & Right(strFilename, Len(strFilename) - 1)
            End If
        End If
    End If
    
    GetSetupPath = strTemp
    Exit Function
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function GetCompressName(ByVal strFilename As String) As String
'����ת��Ϊ7z��׺��ѹ����ʽ����
    On Error GoTo ErrH
    GetCompressName = strFilename & ".7z"
    Exit Function
ErrH:
    If err Then
         MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Sub CompareFile()
'����:�Ƚ��ļ��Ƿ���Ҫ�ռ�
    On Error GoTo ErrH
    
    Dim i      As Long
    Dim strMD5 As String
    Dim lngErr As Long  'δ��װ�ĸ���
    Dim lngSJ  As Long  'δ�ռ��ĸ���
    Dim strFullPath As String
    Dim strFilename As String
    Dim strSetupPath As String
    Dim strFileType As String
    
    If FindFile(mstrSourceFloder) = False Then
        Exit Sub
    End If

    If mshBill.Rows = 0 Then Exit Sub
    For i = 1 To mshBill.Rows - 1
        strMD5 = mshBill.TextMatrix(i, HeadInfor.MD5)
        
        If Len(strMD5) = 0 Then
            
            strFilename = mshBill.TextMatrix(i, HeadInfor.������)
            strSetupPath = mshBill.TextMatrix(i, HeadInfor.��װ·��)
            strFileType = mshBill.TextMatrix(i, HeadInfor.����)
            
            strFullPath = GetSetupPath(Nvl(strFilename, ""), Nvl(strSetupPath, ""), Nvl(strFileType, ""), mstrzlAppSoftPath)
            If FindFile(strFullPath) Then
                mshBill.TextMatrix(i, HeadInfor.��Ϣ) = "δ�ռ����ļ�!"
                mshBill.TextMatrix(i, HeadInfor.�ռ�����) = "1"
                mshBill.SetRowColor i, vbBlue, False
                lngSJ = lngSJ + 1
            Else
                mshBill.TextMatrix(i, HeadInfor.��Ϣ) = "δ��װ�ļ�!"
                mshBill.TextMatrix(i, HeadInfor.�ռ�����) = "2"
                mshBill.SetRowColor i, vbRed, False
                lngErr = lngErr + 1
            End If
            
        End If
    Next
    
    If lngErr = 0 Then
        stbThis.Panels(2).Text = ""
    Else
        stbThis.Panels(2).Text = lngErr & "���ļ�δ��װ " & IIf(lngSJ = 0, "", lngSJ & "���ļ�δ�ռ�")
    End If
    Exit Sub
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Function GetFileName(ByVal strFile As String) As String
'����:ȥ���ļ���׺���ļ���
    Dim i As Integer
    If strFile = "" Then Exit Function
    i = InStrRev(strFile, ".")
    If i > 0 Then
        GetFileName = Left(strFile, i - 1)
    End If
End Function

Private Sub SaveUpList(upList As UpdateList)
    On Error GoTo ErrH
    Dim strSQL As String
    Dim i As Integer
    Dim strFilename As String
    Dim strMD5      As String 'MD5
    Dim str�汾��   As String '�汾��
    Dim str�޸����� As String '�޸�����
    Dim strVision   As String
    Dim strArr    As Variant
    Dim lng�汾��   As Double
    
    If mblnOptType = False Then
        strSQL = "update zlfilesupgrade set MD5= ''"
        gcnOracle.Execute strSQL
    End If
    
    If SafeArrayGetDim(upList.uFile) <> 0 Then
        gcnOracle.BeginTrans
        For i = 0 To UBound(upList.uFile)
            strFilename = upList.uFile(i).FileName
            strMD5 = upList.uFile(i).FileMD5
            str�汾�� = upList.uFile(i).FileVision
            strVision = str�汾��
            If strVision <> "" Then
                strArr = Split(strVision, ".")
                lng�汾�� = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
                strVision = lng�汾��
            End If
            
            str�޸����� = upList.uFile(i).FileEditDate
            
            
            If strFilename <> "" And strMD5 <> "" Then
                strSQL = "update zlfilesupgrade set MD5= '" & strMD5 & "',�汾��='" & strVision & "',�޸�����='" & str�޸����� & "' where upper(�ļ���)='" & UCase(strFilename) & "'"
                gcnOracle.Execute strSQL
            End If
        Next
        gcnOracle.CommitTrans
    End If
    
    Exit Sub
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
        Resume
        gcnOracle.RollbackTrans
    End If
End Sub

Private Function funCanWrite(strWritePath As String) As Boolean
'�ж�Զ��Ŀ¼�Ƿ����дȨ��
    Dim strDest     As String

    On Error GoTo ErrH
            strDest = strWritePath & "\tmp.txt"
            mobjFile.CreateTextFile strDest
            mobjFile.DeleteFile strDest, True
            funCanWrite = True
    Exit Function
ErrH:
    funCanWrite = False
End Function

Private Function GetTmpPath() As String
    Dim tmpBuffer As String
    tmpBuffer = String(255, Chr(0))
    GetTempPath 256, tmpBuffer
    GetTmpPath = Trim(Left(tmpBuffer, InStr(1, tmpBuffer, Chr(0)) - 1))
End Function

Private Sub FloderToClipBoard(ByVal strSourceFloder As String)
    '������ʱ�ռ��ļ�Ŀ¼���ļ�����������ȥ
    Dim strFile() As String
    Dim strSourceFile As String
    Dim strTemp As String
    Dim i As Integer
    strSourceFile = strSourceFloder & "\"
    Erase strFile
    
    
    If mobjFile.FolderExists(strSourceFile) Then
        With FileList
            .Refresh
            .Path = strSourceFile
            .FileName = "*.*"
            
            For i = 0 To .ListCount - 1
                ReDim Preserve strFile(i)
                strTemp = strSourceFile & .List(i)
                strFile(i) = strTemp
            Next
            
            If .ListCount <> 0 Then
                Call clipCopyFiles(strFile)
            End If
        End With
    End If
End Sub

Private Sub BillFileSort()
    On Error GoTo ErrH
    Dim lngRow As Long
    Dim curRow As Long
    Dim strGradeType As String
    curRow = 1
    LoadHeadInforShow
    'strGradeType 0 �����ռ� 1 δ�ռ��ļ� 2 δ��װ·��
    
    '2 Ϊ��װ��·��
    For lngRow = 1 To mshBill.Rows - 1
        strGradeType = mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����)
        If strGradeType = "2" Then
            With mshBillShow
                 .TextMatrix(curRow, HeadInfor.���) = curRow
                 .TextMatrix(curRow, HeadInfor.������) = mshBill.TextMatrix(lngRow, HeadInfor.������)
                 .TextMatrix(curRow, HeadInfor.�汾��) = mshBill.TextMatrix(lngRow, HeadInfor.�汾��)
                 .TextMatrix(curRow, HeadInfor.�޸�����) = mshBill.TextMatrix(lngRow, HeadInfor.�޸�����)
                 .TextMatrix(curRow, HeadInfor.��Ϣ) = mshBill.TextMatrix(lngRow, HeadInfor.��Ϣ)
                 .TextMatrix(curRow, HeadInfor.��������) = mshBill.TextMatrix(lngRow, HeadInfor.��������)
                 .TextMatrix(curRow, HeadInfor.˵��) = mshBill.TextMatrix(lngRow, HeadInfor.˵��)
                 .TextMatrix(curRow, HeadInfor.����) = mshBill.TextMatrix(lngRow, HeadInfor.����)
                 .TextMatrix(curRow, HeadInfor.��װ·��) = mshBill.TextMatrix(lngRow, HeadInfor.��װ·��)
                 .TextMatrix(curRow, HeadInfor.MD5) = mshBill.TextMatrix(lngRow, HeadInfor.MD5)
                 .TextMatrix(curRow, HeadInfor.�ռ�����) = mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����)
                 .Rows = .Rows + 1
                 .SetRowColor curRow, vbRed, False
            End With
            curRow = curRow + 1
        End If
    Next
    
    '1 δ�ռ����ļ�
    For lngRow = 1 To mshBill.Rows - 1
        strGradeType = mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����)
        If strGradeType = "1" Then
            With mshBillShow
                 .TextMatrix(curRow, HeadInfor.���) = curRow
                 .TextMatrix(curRow, HeadInfor.������) = mshBill.TextMatrix(lngRow, HeadInfor.������)
                 .TextMatrix(curRow, HeadInfor.�汾��) = mshBill.TextMatrix(lngRow, HeadInfor.�汾��)
                 .TextMatrix(curRow, HeadInfor.�޸�����) = mshBill.TextMatrix(lngRow, HeadInfor.�޸�����)
                 .TextMatrix(curRow, HeadInfor.��Ϣ) = mshBill.TextMatrix(lngRow, HeadInfor.��Ϣ)
                 .TextMatrix(curRow, HeadInfor.��������) = mshBill.TextMatrix(lngRow, HeadInfor.��������)
                 .TextMatrix(curRow, HeadInfor.˵��) = mshBill.TextMatrix(lngRow, HeadInfor.˵��)
                 .TextMatrix(curRow, HeadInfor.����) = mshBill.TextMatrix(lngRow, HeadInfor.����)
                 .TextMatrix(curRow, HeadInfor.��װ·��) = mshBill.TextMatrix(lngRow, HeadInfor.��װ·��)
                 .TextMatrix(curRow, HeadInfor.MD5) = mshBill.TextMatrix(lngRow, HeadInfor.MD5)
                 .TextMatrix(curRow, HeadInfor.�ռ�����) = mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����)
                 .Rows = .Rows + 1
                 .SetRowColor curRow, vbBlue, False
            End With
            curRow = curRow + 1
        End If
    Next
    
    '0 δ�ռ����ļ�
    For lngRow = 1 To mshBill.Rows - 1
        strGradeType = mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����)
        If strGradeType = "0" Then
            With mshBillShow
                 .TextMatrix(curRow, HeadInfor.���) = curRow
                 .TextMatrix(curRow, HeadInfor.������) = mshBill.TextMatrix(lngRow, HeadInfor.������)
                 .TextMatrix(curRow, HeadInfor.�汾��) = mshBill.TextMatrix(lngRow, HeadInfor.�汾��)
                 .TextMatrix(curRow, HeadInfor.�޸�����) = mshBill.TextMatrix(lngRow, HeadInfor.�޸�����)
                 .TextMatrix(curRow, HeadInfor.��Ϣ) = mshBill.TextMatrix(lngRow, HeadInfor.��Ϣ)
                 .TextMatrix(curRow, HeadInfor.��������) = mshBill.TextMatrix(lngRow, HeadInfor.��������)
                 .TextMatrix(curRow, HeadInfor.˵��) = mshBill.TextMatrix(lngRow, HeadInfor.˵��)
                 .TextMatrix(curRow, HeadInfor.����) = mshBill.TextMatrix(lngRow, HeadInfor.����)
                 .TextMatrix(curRow, HeadInfor.��װ·��) = mshBill.TextMatrix(lngRow, HeadInfor.��װ·��)
                 .TextMatrix(curRow, HeadInfor.MD5) = mshBill.TextMatrix(lngRow, HeadInfor.MD5)
                 .TextMatrix(curRow, HeadInfor.�ռ�����) = mshBill.TextMatrix(lngRow, HeadInfor.�ռ�����)
                 .Rows = .Rows + 1
                 .SetRowColor curRow, &HFFFFFF, False
            End With
            curRow = curRow + 1
        End If
    Next
    
    If mshBillShow.Rows <> 0 Then
        mshBillShow.Rows = mshBillShow.Rows - 1
    End If
    
    Me.mshBill.Visible = False
    Me.mshBillShow.Visible = True
    Exit Sub
ErrH:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub chk��������_Click()
    If chk��������.value = 1 Then
        DTP��������.Enabled = True
        cmd��������.Enabled = True
    Else
        DTP��������.Enabled = False
        cmd��������.Enabled = True
    End If
End Sub

Private Sub cmd��������_Click()
    Dim strNow As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ErrHand
    
    If chk��������.value = 1 Then
        strNow = Format(CurrentDate(), "yyyy-MM-dd")
        If DTP�������� < CDate(strNow) Then
            MsgBox "�������ڲ���С�ڵ�ǰ����������!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        
        Set rsTmp = New ADODB.Recordset
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '�ͻ�����������'"
        Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Update zlRegInfo Set ����='" & Format(DTP��������, "yyyy-MM-dd") & "' Where ��Ŀ='�ͻ�����������'"
            gcnOracle.Execute strSQL
        Else
            strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ͻ�����������',Null,'" & Format(DTP��������, "yyyy-MM-dd") & "')"
            gcnOracle.Execute strSQL
        End If
        
    Else
        Set rsTmp = New ADODB.Recordset
        gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '�ͻ�����������'"
        Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Update zlRegInfo Set ����=Null Where ��Ŀ='�ͻ�����������'"
            gcnOracle.Execute strSQL
        Else
            strSQL = "INSERT INTO zlRegInfo (��Ŀ,�к�,����) VALUES ('�ͻ�����������',Null,Null)"
            gcnOracle.Execute strSQL
        End If
    
    End If

  Exit Sub
ErrHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub OpinionUpGradeDate()
    '�ж��Ƿ��趨����������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
     
    Set rsTmp = New ADODB.Recordset
    gstrSQL = "Select ��Ŀ,���� From zlRegInfo where ��Ŀ = '�ͻ�����������'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    cmd��������.Enabled = True
    
    If rsTmp.EOF = False Then
        If Nvl(rsTmp!����) = "" Then
            chk��������.Enabled = True
            DTP��������.Enabled = False
            
            chk��������.value = 0
            DTP��������.value = Format(CurrentDate(), "yyyy-MM-dd")
        Else
            chk��������.Enabled = True
            DTP��������.Enabled = True
        
            chk��������.value = 1
            DTP��������.value = Nvl(rsTmp!����, Format(CurrentDate(), "yyyy-MM-dd"))
        End If
    Else
        chk��������.Enabled = True
        DTP��������.Enabled = False
        
        chk��������.value = 0
        DTP��������.value = Format(CurrentDate(), "yyyy-MM-dd")
    End If
    
    Exit Sub
ErrHand:
End Sub
